'No Good Gofers by Bodydump
'Models by Dark
Option Explicit
Randomize

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Const BallSize = 50
Const BallMass = 1

Dim SideFlashersEnabled
Dim UseVPMDMD
UseVPMDMD = True
Dim VRCab_BladesEnabled
VRCab_BladesEnabled = 1  ' Default to enabled
Dim apronWallStyle : apronWallStyle = 1  ' 1 = Style A, 2 = Style B
Dim sidebladeStyle : sidebladeStyle = 1  ' 1 = Normal, 2 = Artwork
Dim GIrightEnabled, GIleftEnabled
Dim GIEnabled
Dim BackTopEnabled
Dim Table1
Dim tablewidth: tablewidth = Table.width
Dim tableheight: tableheight = Table.height
Const tnob = 7

Dim LF : Set LF = New FlipperPolarity
Dim RF : Set RF = New FlipperPolarity

'Const bladeArt = 1 '1=On (Art), 2=Alt, 0=Sideblades Off.    'removed

'***Apron Style***
' 0 = Standard apron
' 1 = Custom apron
apronstyle = 0

'**************************
'Desktop/Fullscreen Changes
'**************************

'If Table.ShowDT = true then
' Ramp15.visible = True
' Ramp16.visible = True
' Ramp17.visible = True
'else
' Ramp15.visible = False
' Ramp16.visible = False
' Ramp17.visible = False
'end if


   LoadVPM"01520000","WPC.VBS",3.1


'********************
'Standard definitions
'********************

    Dim xx, apronstyle
    Dim Bump1, Bump2, Bump3, RGoferIsUp, RightRampUp, LGoferIsUp, LeftRampUp
  Dim bsTrough, bsLeftEject, bsPuttOutPopper, bsJetPopper, bsUpperRightEject, mTT, cbCaptive
  LGoferIsUp = 0
  LeftRampUp = 0
  RGoferIsUp = 0
  RightRampUp = 0
  Const cGameName = "ngg_13"
  Const UseSolenoids = 2
  Const UseLamps = 0
  Const UseGI = 0
  Const UseSync = 0
  Const HandleMechs = 1

  Set GICallback = GetRef("UpdateGIon")
  Set GICallback2 = GetRef("UpdateGI")

  Const SSolenoidOn="fx_Solenoid"
  Const SSolenoidOff="fx_solenoidoff"
  Const sCoin="fx_Coin"

'***********
' Solenoids
' the commented solenoids are not in used in this script
'***********

    SolCallback(1)="vpmSolAutoPlunger Plunger1,3,"          'Autofire
    SolCallback(2)="vpmSolAutoPlunger KickBack,3,"          'KickBack
    'SolCallback(2) = "SolKickbackWrapper"
    SolCallback(3)="SolPuttOut"                   'Clubhouse Kicker
    SolCallback(4)="SolLeftGoferUP"                 'Left Gofer Up
    SolCallback(5)="SolRightGoferUP"                'Right Gofer Up
    SolCallback(6)="bsJetPopper.SolOut"               'Jet Popper
    SolCallback(7)="bsLeftEject.SolOut"               'Left Eject (Sand Trap)
    SolCallback(8)="bsUpperRightEject.SolOut"           'Upper Right Eject
    SolCallback(9)="SolBallRelease"                 'Trough Eject
'   'SolCallback(10)="vpmSolSound ""Sling"","           'Left Slingshot
'   'SolCallback(11)="vpmSolSound ""Sling"","           'Right Slingshot
'   'SolCallback(12)="vpmSolSound ""Jet3"","            'Top Jet Bumper
'   'SolCallback(13)="vpmSolSound ""Jet3"","            'Middle Jet Bumper
'   'SolCallback(14)="vpmSolSound ""Jet3"","            'Bottom Jet Bumper
    SolCallback(15)="SolLeftGoferDown"                'Left Gofer Down
    SolCallback(16)="SolRightGoferDown"               'Right Gofer Down
    SolCallback(17)="SetLamp 170,"                  'Jet Flasher
    SolCallback(18)="SetLamp 178,"                  'Lower Left Flasher
    SolCallback(19)="SetLamp 180,"                  'Left Spinner Flasher
    SolCallback(20)="SetLamp 122,"                  'Right Spinner Flasher
    SolCallback(21)="SetLamp 179,"                  'Lower Right Flasher
    SolCallback(24)="SolSubway"                   'Underground Pass
    SolCallback(25)="SetLamp 131,"                  'Sand Trap Flasher
    SolCallback(26)="SetLamp 190,"                  'Wheel Flasher
    SolCallback(27)="SolLeftRampDown"               'Left Ramp Down
    SolCallback(28)="SolRightRampDown"                'Right Ramp Down
    SolCallback(35)="SolSlamRamp"                 'Ball Launch Ramp
    SolCallback(51)="SetLamp 177,"                  'Upper Right Flasher 1
    SolCallback(52)="SetLamp 175,"                  'Upper Right Flasher 2
    SolCallback(53)="SetLamp 173,"                  'Upper Right Flasher 3
    SolCallback(56)="SetLamp 176,"                  'Upper Left Flasher 1
    SolCallback(57)="SetLamp 174,"                  'Upper Left Flasher 2
    SolCallback(58)="SetLamp 172,"                  'Upper Left Flasher 3
'   SolCallback(sLRFlipper)="vpmSolFlipper RightFlipper,Nothing,"
'   SolCallback(sLLFlipper)="vpmSolFlipper LeftFlipper,Nothing,"
'   SolCallback(sURFlipper)="vpmSolFlipper UpperFlipper,Nothing,"
    SolCallback(54)="SetLamp 145,"
    SolCallback(55)="SetLamp 146,"

'Sub SolKickbackWrapper(Enabled)
'    If Enabled Then
'        ' --- your custom sound here ---
' PlaySoundAtLevelStatic ("WarehouseKick"), DrainSoundLevel, Kickback
'
'       ' --- then let VPX perform the kick exactly as before ---
'        vpmSolAutoPlunger KickBack, 10, Enabled
'    End If
'End Sub

'
  Sub SolBallRelease(Enabled)
      If Enabled Then
        If bsTrough.Balls Then
          vpmTimer.PulseSw 31
                    RandomSoundBallRelease BallRelease
        End If
      bsTrough.ExitSol_On
      End If
  End Sub

  Sub SolPuttOut(Enabled)
    If Enabled Then
      If bsPuttOutPopper.Balls Then
        bsPuttOutPopper.InitKick sw44a,170,12
        PlaySound"scoopexit"
      End If
      bsPuttOutPopper.ExitSol_On
    End If
  End Sub
  Sub SolSubway(Enabled)
    If Enabled Then
      If bsPuttOutPopper.Balls Then
        bsPuttOutPopper.InitKick subwaystart,180,3
        bsPuttOutPopper.ExitSol_On
      End If
    End If
  End Sub

  Sub SolLeftGoferUP(Enabled)
    If Enabled Then
      LrampTimer.Enabled = 0
      If LGoferIsUp = 0 Then
        Controller.Switch(41)=0
        Controller.Switch(47)=0
        LGofer.IsDropped=0
        LGoferUp.Enabled = 1
        LGoferDown.Enabled = 0
        Lramp.collidable=0
        Lramp1.collidable=1
        LeftRampUp = 1
      Else
        Controller.Switch(41)=0
        Controller.Switch(47)=0
        LGoferUp.Enabled = 1
        LeftGoferBlock.collidable = 1
        LGPos = 6
      End If
    End If
  End Sub

  Sub SolRightGoferUP(Enabled)
    If Enabled Then
      RrampTimer.Enabled = 0
      If RGoferIsUp = 0 Then
        Controller.Switch(42)=0
        Controller.Switch(48)=0
        RGofer.IsDropped=0
        RGoferUp.Enabled = 1
        RGoferDown.Enabled = 0
        Rramp.collidable=0
        Rramp1.collidable=1
        RightRampUp = 1
      Else
        Controller.Switch(42)=0
        Controller.Switch(48)=0
        RGoferUp.Enabled = 1
        RGPos = 6
      End If
    End If
  End Sub

  Sub SolLeftGoferDown(Enabled)
    If Enabled Then
      Controller.Switch(41)=1
      LGoferDown.Enabled = 1
      LGoferUp.Enabled = 0
      LGofer.IsDropped=1
    End If
  End Sub

  Sub SolRightGoferDown(Enabled)
    If Enabled Then
      Controller.Switch(42)=1
      RGoferDown.Enabled = 1
      RGoferUp.Enabled = 0
      RGofer.IsDropped = 1
    End If
  End Sub

  Sub SolLeftRampDown(Enabled)
    If Enabled Then
      RightRampUp = 0
      LrampDir = 1
      Lramp.collidable=1
      Lramp1.collidable=0
      PlaySound "gate"
      LrampTimer.Enabled = 1

    Else

    End If
  End Sub

  Sub SolRightRampDown(Enabled)
    If Enabled Then
      RightRampUp = 0
      RrampDir = 1
      Rramp.collidable=1
      Rramp1.collidable=0
      PlaySound "gate"
      RrampTimer.Enabled = 1

    Else

    End If
  End Sub

  Sub SolSlamRamp(Enabled)
    If Enabled Then
      For each xx in aslamramp:xx.IsDropped=0:Next
      For each xx in aslamrampup:xx.IsDropped=1:Next
      jumpramptrigger.enabled = 1
      jumpramp.Collidable=1
      slamrampup.Collidable = 0
      Flipper1.RotateToEnd
      rampshadow.visible = 0
      SlamRampDownTimer.Enabled = 1
    Else

    End If
  End Sub

Sub SlamRampDownTimer_Timer()
    SlamRampUpTimer.Enabled = 1
    SlamRampDownTimer.Enabled = 0
End Sub
Sub SlamRampUpTimer_Timer()
    For each xx in aslamramp:xx.IsDropped=1:Next
    For each xx in aslamrampup:xx.IsDropped=9:Next
    jumpramp.Collidable=0
    slamrampup.Collidable = 1
    jumpramptrigger.enabled = 0
    Flipper1.RotateToStart
    rampshadow.visible = 1
    SlamRampUpTimer.Enabled = 0
End Sub

Sub UpdateBackTopVisibility()
    On Error Resume Next

    Dim shouldShow
    shouldShow = (Table.ShowDT = False And VRRoom = 0 And BackTopEnabled = 1)

    If shouldShow Then
        Back_top.Opacity = 100
        Back_top.Image = "captop"
    Else
        Back_top.Opacity = 0
        Back_top.Image = ""
    End If

    On Error Goto 0
End Sub

Sub Table_ShowDTToggle(ByVal isDT)
    vpmTimer.AddTimer 100, "UpdateBackTopVisibility"
    vpmTimer.AddTimer 150, "LoadVRRoom" ' Refresh VR room settings when toggling DT
End Sub

Sub UpdateSideFlashers()
    On Error Resume Next

    Dim visState
    If SideFlashersEnabled Then
        visState = 1
    Else
        visState = 0
    End If

    ' Objects without spaces
    Red_Pumper.visible = visState
    Blue_Pumper.visible = visState
    White_Pumper.visible = visState

    ' Objects WITH spaces - use square brackets
    [L178 Side Flasher].visible = visState
    [L180 Side Flasher].visible = visState
    [L179 Side Flasher].visible = visState

    On Error Goto 0
End Sub

'************
' Table init.
'************

  Sub Table_Init
    InitPolarity
  vpmInit me
  With Controller
    .GameName = cGameName
    If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description:Exit Sub
    .SplashInfoLine="No Good Gofers" & vbNewLine & "VP Table by Bodydump"
    .HandleKeyboard=0
    .ShowTitle=0
    .ShowDMDOnly=1
    .ShowFrame=0
    .HandleMechanics=1
    .Hidden=0
    On Error Resume Next
    .Run GetPlayerHWnd
    If Err Then MsgBox Err.Description
    On Error Goto 0
  End With
  'Nudging
      vpmNudge.TiltSwitch=14
      vpmNudge.Sensitivity=5
      vpmNudge.TiltObj=Array(Bumper1,Bumper2,Bumper3,LeftSlingshot,RightSlingshot)

  'Trough
    Set bsTrough=New cvpmBallStack
    With bsTrough
      .InitSw 0,32,33,34,35,36,37,0
      .InitKick BallRelease,105,5
      .InitEntrySnd "fx_Solenoid", "fx_Solenoid"
      .InitExitSnd SoundFX("",DOFContactors), SoundFX("fx_Solenoid",DOFContactors)
      .Balls=6
    End With

  'Sand Trap Eject
    Set bsLeftEject=New cvpmBallStack
    With bsLeftEject
      .InitSaucer LeftEject,78,78,25
      .KickForceVar = 3
      .KickAngleVar = 3
      .InitExitSnd SoundFX("Saucer_Kick",DOFContactors), SoundFX("fx_Solenoid",DOFContactors)
    End With

  'Putt Out Popper
    Set bsPuttOutPopper=New cvpmBallStack
    With bsPuttOutPopper
      .InitSw 0,44,0,0,0,0,0,0
      .InitKick sw44a,170,12
      .KickForceVar = 3
      .KickAngleVar = 3
    End With

  'Jet Popper
    Set bsJetPopper=New cvpmBallStack
    With bsJetPopper
      .InitSw 0,38,25,0,0,0,0,0
      .InitKick JetPopper,180,10
      .KickForceVar = 3
      .KickAngleVar = 3
      .InitExitSnd SoundFX("fx_Popper",DOFContactors), SoundFX("fx_Solenoid",DOFContactors)
    End With

  'Upper Right Eject
    Set bsUpperRightEject=New cvpmBallStack
    With bsUpperRightEject
      .InitSw 0,46,0,0,0,0,0,0
      .InitKick upperrightkicker,120,5
      .KickForceVar = 3
      .KickAngleVar = 3
      .InitExitSnd SoundFX("fx_Popper",DOFContactors), SoundFX("fx_Solenoid",DOFContactors)
    End With

  'Turntable
    Set mTT=New cvpmTurnTable
    With mTT
      .InitTurnTable TurnTable,40
      .SpinUp=40
      .SpinDown=40
      .CreateEvents "mTT"
    End With

  'Captive Ball
    Set cbCaptive=New cvpmCaptiveBall
    With cbCaptive
      .InitCaptive Captive,CWALL,Array(Captive1,Captive2),350
      .NailedBalls=1
      .Start
      .ForceTrans=0.9
      .MinForce=4
      .CreateEvents "cbCaptive"
    End With
    Captive1.CreateBall
  If apronstyle = 0 Then
    apron2.IsDropped = 1
    apron1.IsDropped = 0
  Else
    apron2.IsDropped = 0
    apron1.IsDropped = 1
  End If

  If nightday<5 then
  For each xx in extraambient:xx.intensityscale = 1:Next
  End If
  If nightday>5 then
  For each xx in extraambient:xx.intensity = xx.intensity*.9:Next
  End If
  If nightday>10 then
  For each xx in extraambient:xx.intensity = xx.intensity*.7:Next
  For each xx in GIleft:xx.intensity = xx.intensity*.9:Next
  For each xx in GIright:xx.intensity = xx.intensity*.9:Next
  End If
  If nightday>20 then
  For each xx in extraambient:xx.intensity = xx.intensity*.6:Next
  For each xx in GIleft:xx.intensity = xx.intensity*.8:Next
  For each xx in GIright:xx.intensity = xx.intensity*.8:Next
  End If
  If nightday>30 then
  For each xx in extraambient:xx.intensity = xx.intensity*.4:Next
  For each xx in GIleft:xx.intensity = xx.intensity*.7:Next
  For each xx in GIright:xx.intensity = xx.intensity*.7:Next
  End If



'      '**Main Timer init
           PinMAMETimer.Enabled = 1

  Plunger1.PullBack
  Kickback.PullBack
  Controller.Switch(22)=1 'close coin door
  Controller.Switch(24)=0 'always closed
  Controller.Switch(41)=1 'drop left gofer
  Controller.Switch(42)=1 'drop right gofer
  Controller.Switch(47)=1 'left ramp down
  Controller.Switch(48)=1 'right ramp down
  LGofer.IsDropped = 1
  RGofer.IsDropped = 1
  Rramp1.collidable=0
  Lramp1.collidable=0

   LockbarEnabled = 1 ' Default to On
    BallType = 2 ' Default to Original ball
    BackTopEnabled = 1
    UpdateBackTopVisibility
    LastBallType = 2 ' Initialize the tracking variable
    BrightnessFadeSpeed = 0.03 ' Adjust this for faster/slower fade
  CurrentBrightness = 1.0 ' Start at full brightness
  TargetBrightness = 1.0
    ' Table.Brightness is not valid - we'll adjust light intensity instead
   GIrightEnabled = True
  GIleftEnabled = True
  sidebladeStyle = 1  ' Default to Normal style
  UpdateBladeVisibility
    apronWallStyle = 1  ' Default to Style 1
    UpdateApronWall ' Set initial apron wall texture
    SideFlashersEnabled = 1 ' Default to ON
    UpdateSideFlashers ' Set initial visibility


   Table.BallImage = "old_ass_eyes_ball"



End Sub

' Choose Side Blades    'removed
' if bladeArt = 1 then
'   PinCab_Blades.Image = "Sidewalls NGG"
'   PinCab_Blades.visible = 1
'    elseif bladeArt = 2 then
'   PinCab_Blades.Image = "Sidewalls NGG2"
'   PinCab_Blades.visible = 1
' elseif bladeArt = 0 then
'   PinCab_Blades.visible = 0
' End if




    Sub Table_Paused:Controller.Pause = 1:End Sub
    Sub Table_unPaused:Controller.Pause = 0:End Sub


'**********
' Keys
'**********

' Plunger animation timers
Sub TimerPlunger_Timer

  If VRCab_plunger.Y < -115 then
      VRCab_plunger.Y = VRCab_plunger.Y + 5
  End If
End Sub

Sub TimerPlunger2_Timer
  VRCAb_plunger.Y = -235.78 + (5* Plunger.Position) -20
End Sub


    Sub table_KeyDown(ByVal Keycode)

      If keycode = StartGameKey Then
        If VRRoom > 0 or VRTest Then
          VRCab_StartButton.y = VRCab_StartButton.y - 2
          VRCab_StartButtonRim.y = VRCab_StartButtonRim.y - 2
        End If
      End If

      If keycode = RightFlipperKey Then
        FlipperActivate RightFlipper, RFPress
        If VRRoom > 0 or VRTest Then VRCab_FlipperButtonRight.x = VRCab_FlipperButtonRight.x - 5
      End If
      If keycode = LeftFlipperKey Then
        FlipperActivate LeftFlipper, LFPress
        If VRRoom > 0 or VRTest Then VRCab_FlipperButtonLeft.x = VRCab_FlipperButtonLeft.x + 5
      End If

      If keycode = PlungerKey Then
        TimerPlunger.Enabled = True
        TimerPlunger2.Enabled = False
      End If

      If keycode = PlungerKey Then Plunger.Pullback:PlaySound SoundFX("fx_plungerpull",DOFContactors)
      If keycode = LeftTiltKey Then Nudge 90, 1:SoundNudgeLeft
      If keycode = RightTiltKey Then Nudge 270, 1:SoundNudgeRight
      If keycode = CenterTiltKey Then Nudge 0, 1:SoundNudgeCenter
            If keycode = LeftFlipperKey Then FlipperActivate LeftFlipper, LFPress
            If keycode = RightFlipperKey Then FlipperActivate RightFlipper, RFPress
      If vpmKeyDown(keycode) Then Exit Sub

    End Sub

    Sub table_KeyUp(ByVal Keycode)
  If keycode = AddCreditKey Then
    Select Case Int(Rnd * 3)
      Case 0: PlaySound "Coin_In_1", 0, CoinSoundLevel, 0, 0.25
      Case 1: PlaySound "Coin_In_2", 0, CoinSoundLevel, 0, 0.25
      Case 2: PlaySound "Coin_In_3", 0, CoinSoundLevel, 0, 0.25
    End Select
  End If

      If keycode = PlungerKey Then
        TimerPlunger.Enabled = False
        TimerPlunger2.Enabled = True
        VRCab_Plunger.Y = -235.78
      end if

      If keycode = StartGameKey Then
        If VRRoom > 0 or VRTest Then
          VRCab_StartButton.y = VRCab_StartButton.y + 2
          VRCab_StartButtonRim.y = VRCab_StartButtonRim.y + 2
        End If
      End If

      If keycode = RightFlipperKey Then
        FlipperDeActivate RightFlipper, RFPress
        If VRRoom > 0  or VRTest Then VRCab_FlipperButtonRight.x = 2102
      End If

      If keycode = LeftFlipperKey Then
        FlipperDeActivate LeftFlipper, LFPress
        If VRRoom > 0 or VRTest Then VRCab_FlipperButtonLeft.x = 2099
      End If


      'If keycode = PlungerKey Then Plunger.Fire:PlaySound SoundFX("fx_plunger",DOFContactors)
            If keycode = PlungerKey Then Plunger.Fire:SoundPlungerReleaseBall()
            If keycode = LeftFlipperKey Then FlipperDeActivate LeftFlipper, LFPress
            If keycode = RightFlipperKey Then FlipperDeActivate RightFlipper, RFPress
      If vpmKeyUp(keycode) Then Exit Sub
    End Sub


'********************
'    Flippers
'********************
Const ReflipAngle = 20


    SolCallback(sLRFlipper) = "SolRFlipper"
    SolCallback(sURFlipper) = "SolURFlipper"
    SolCallback(sLLFlipper) = "SolLFlipper"

    Sub SolLFlipper(Enabled)
      If Enabled Then
        LF.Fire
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

    Sub SolRFlipper(Enabled)
      If Enabled Then
  '             UpperFlipper.RotateToEnd
        RF.Fire
    If rightflipper.currentangle > rightflipper.endangle - ReflipAngle Then
      RandomSoundReflipUpRight RightFlipper
    Else
      SoundFlipperUpAttackRight RightFlipper
      RandomSoundFlipperUpRight RightFlipper
    End If
  Else
    RightFlipper.RotateToStart
 '       UpperFlipper.RotateToStart
    If RightFlipper.currentangle > RightFlipper.startAngle + 5 Then
      RandomSoundFlipperDownRight RightFlipper
    End If
    FlipperRightHitParm = FlipperUpSoundLevel
  End If
End Sub

Sub SolURFlipper(Enabled)
    If Enabled Then
        UpperFlipper.RotateToEnd
        If upperflipper.currentangle < upperflipper.endangle + ReflipAngle Then
            RandomSoundReflipUpRight upperflipper
        Else
            SoundFlipperUpAttackRight upperflipper
      RandomSoundFlipperUpRight upperflipper
        End If
    Else
        upperflipper.RotateToStart
        If RightFlipper.currentangle > RightFlipper.startAngle + 5 Then
      RandomSoundFlipperDownRight RightFlipper
    End If
    FlipperRightHitParm = FlipperUpSoundLevel
    End If
End Sub


'****************************
' Drain holes, vuks & saucers
'****************************
    Sub Drain_Hit:RandomSoundDrain drain:bsTrough.AddBall Me:End Sub
    Sub LeftEject_Hit:PlaySound "Saucer_Enter_2", 0, 1, -0.05, 0.05:bsLeftEject.AddBall 0:End Sub     'Sand Trap Eject
    Sub sw44a_Hit:RandomSoundDelayedBallDropOnPlayfield sw44a:bsPuttOutPopper.AddBall Me:End Sub      'Putt Out Popper
    Sub sw44b_Hit:RandomSoundDelayedBallDropOnPlayfield sw44b:bsPuttOutPopper.AddBall Me:End Sub      'Putt Out Popper
    Sub sw44c_Hit:RandomSoundDelayedBallDropOnPlayfield sw44c:bsPuttOutPopper.AddBall Me:End Sub      'Putt Out Popper
    Sub sw44d_Hit:RandomSoundDelayedBallDropOnPlayfield sw44d:bsPuttOutPopper.AddBall Me:End Sub      'Putt Out Popper
    Sub sw44e_Hit:RandomSoundDelayedBallDropOnPlayfield sw44e:bsPuttOutPopper.AddBall Me:End Sub      'Putt Out Popper
    Sub sw44f_Hit:RandomSoundDelayedBallDropOnPlayfield sw44f:bsPuttOutPopper.AddBall Me:End Sub      'Putt Out Popper
    Sub sw44g_Hit:RandomSoundDelayedBallDropOnPlayfield sw44g:bsPuttOutPopper.AddBall Me:End Sub      'Putt Out Popper
    Sub sw44h_Hit:RandomSoundDelayedBallDropOnPlayfield sw44h:bsPuttOutPopper.AddBall Me:End Sub      'Putt Out Popper
    Sub sw44i_Hit:RandomSoundDelayedBallDropOnPlayfield sw44i:bsPuttOutPopper.AddBall Me:End Sub      'Putt Out Popper
    Sub subwayend_Hit:bsJetPopper.AddBall Me:End Sub                'Underground Pass
    Sub holeinone_Hit:RandomSoundDelayedBallDropOnPlayfield holeinone:Me.DestroyBall:vpmTimer.PulseSwitch(68),160,"HandleTopHole":End Sub
    Sub behindleftgofer_Hit:PlaySound "fx_kicker_enter", 0, 1, -0.05, 0.05:Me.DestroyBall:vpmTimer.PulseSwitch(67),100,"HandlePuttOut":End Sub
    Sub rightpopperjam_Hit:PlaySound "fx_kicker_enter", 0, 1, 0.05, 0.05:Me.DestroyBall:vpmTimer.PulseSwitch(45),160,"HandleRightGofer":End Sub
    Sub HandleTopHole(swNo):vpmTimer.PulseSwitch(67),100,"HandlePuttOut":End Sub
    Sub HandlePuttOut(swNo):bsPuttOutPopper.AddBall 0:End Sub
    Sub HandleRightGofer(swNo):bsUpperRightEject.AddBall 0:End Sub

'***************
'  Slingshots
'***************
'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
    Dim RStep, Lstep

    Sub RightSlingShot_Slingshot
            RS.VelocityCorrect(Activeball)
      RandomSoundSlingshotRight SLING1
      vpmTimer.PulseSw 52
      RSling.Visible = 0
      RSling1.Visible = 1
      sling1.TransZ = -20
      RStep = 0
      RightSlingShot.TimerEnabled = 1
    End Sub

    Sub RightSlingShot_Timer
      Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.TransZ = -10
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.TransZ = 0:RightSlingShot.TimerEnabled = 0
      End Select
      RStep = RStep + 1
    End Sub

    Sub LeftSlingShot_Slingshot
            LS.VelocityCorrect(Activeball)
      RandomSoundSlingshotLeft SLING2
      vpmTimer.PulseSw 51
      LSling.Visible = 0
      LSling1.Visible = 1
      sling2.TransZ = -20
      LStep = 0
      LeftSlingShot.TimerEnabled = 1
    End Sub

    Sub LeftSlingShot_Timer
      Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.TransZ = -10
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.TransZ = 0:LeftSlingShot.TimerEnabled = 0
      End Select
      LStep = LStep + 1
    End Sub

'***************
'   Bumpers
'***************
    Sub Bumper1_Hit:vpmTimer.PulseSw 54:RandomSoundBumperMiddle Bumper1:End Sub
    Sub Bumper2_Hit:vpmTimer.PulseSw 53:RandomSoundBumperTop Bumper2:End Sub
    Sub Bumper3_Hit:vpmTimer.PulseSw 55:RandomSoundBumperBottom Bumper3:End Sub


'************
' Spinners
'************
    Sub Spinner1_Spin():vpmTimer.PulseSwitch(61),0,"":SoundSpinner Spinner1: End Sub
    Sub Spinner2_Spin():vpmTimer.PulseSwitch(62),0,"":SoundSpinner Spinner2: End Sub

'*********************
' Switches & Rollovers
'*********************
    Sub sw18_Hit:Controller.Switch(18)=1:PlaySound "fx_sensor":Light002.state = 2:End Sub
    Sub sw18_UnHit:Controller.Switch(18)=0:Light002.state = 0:End Sub
    Sub sw16_Hit:Controller.Switch(16)=1:PlaySound "fx_sensor":End Sub
    Sub sw16_UnHit:Controller.Switch(16)=0:End Sub
    Sub sw17_Hit:Controller.Switch(17)=1:PlaySound "fx_sensor":End Sub
    Sub sw17_UnHit:Controller.Switch(17)=0:End Sub
    Sub sw26_Hit:Controller.Switch(26)=1:PlaySound "fx_sensor":End Sub
    Sub sw26_UnHit:Controller.Switch(26)=0:End Sub
    Sub sw27_Hit:Controller.Switch(27)=1:PlaySound "fx_sensor":End Sub
    Sub sw27_UnHit:Controller.Switch(27)=0:End Sub
    Sub sw28_Hit:Controller.Switch(28)=1:PlaySound "fx_sensor":End Sub
    Sub sw28_UnHit:Controller.Switch(28)=0:End Sub
    Sub sw71_Hit:Controller.Switch(71)=1:PlaySound "fx_sensor":End Sub
    Sub sw71_UnHit:Controller.Switch(71)=0:End Sub
    Sub sw72_Hit:Controller.Switch(72)=1:PlaySound "fx_sensor":End Sub
    Sub sw72_UnHit:Controller.Switch(72)=0:End Sub
    Sub sw86_Hit:Controller.Switch(86)=1:PlaySound "Saucer_Enter_2":End Sub
    Sub sw86_UnHit:Controller.Switch(86)=0:End Sub

'***************
'  Targets
'***************

    Sub sw23_Hit
      vpmTimer.PulseSw 23
      'PlaySound SoundFX("fx_target",DOFContactors), 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0
    End Sub

    Sub sw56_Hit
      vpmTimer.PulseSw 56
      'PlaySound SoundFX("fx_target",DOFContactors), 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0
    End Sub

    Sub sw57_Hit
      vpmTimer.PulseSw 57
      'PlaySound SoundFX("fx_target",DOFContactors), 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0
    End Sub

    Sub sw58_Hit
      vpmTimer.PulseSw 58
      'PlaySound SoundFX("fx_target",DOFContactors), 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0
    End Sub

    Sub sw77_Hit
      vpmTimer.PulseSw 77
      'PlaySound SoundFX("fx_target",DOFContactors), 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0
    End Sub

    Sub sw81_Hit
      vpmTimer.PulseSw 81
      'PlaySound SoundFX("fx_target",DOFContactors), 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0
    End Sub

    Sub sw82_Hit
      vpmTimer.PulseSw 82
      'PlaySound SoundFX("fx_target",DOFContactors), 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0
    End Sub

    Sub sw83_Hit
      vpmTimer.PulseSw 83
      'PlaySound SoundFX("fx_target",DOFContactors), 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0
    End Sub

    Sub sw84_Hit
      vpmTimer.PulseSw 84
      'PlaySound SoundFX("fx_target",DOFContactors), 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0
    End Sub

    Sub sw85_Hit
      vpmTimer.PulseSw 85
      'PlaySound SoundFX("fx_target",DOFContactors), 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0
    End Sub

'*********************
' Random Switches
'*********************
    Sub Gate4_Hit():vpmTimer.PulseSw 12:End Sub
    Sub sw15_Hit():vpmTimer.PulseSw 15:End Sub
    Sub cart_Hit():PlaySound "plastichit":vpmTimer.PulseSw 74:GolfCart.Y = 218:WheelTargetPosition.Y = 217:Me.TimerEnabled = 1:End Sub      'Golf Cart
    Sub cart_Timer:GolfCart.Y = 215:WheelTargetPosition.Y = 211:Me.TimerEnabled = 0:End Sub
    Sub sw73_Hit():vpmTimer.PulseSw 73:End Sub
    Sub LGofer_Hit():PlaySound "plastichit":vpmTimer.PulseSw 65:LeftGoferBlock.Y = 441:Me.TimerEnabled = 1:End Sub
    Sub LGofer_Timer:LeftGoferBlock.Y = 449:Me.TimerEnabled = 0:End Sub
    Sub RGofer_Hit():PlaySound "plastichit":vpmTimer.PulseSw 75:RightGoferBlock.Y = 516.625:Me.TimerEnabled = 1:End Sub
    Sub RGofer_Timer:RightGoferBlock.Y = 524.625:Me.TimerEnabled = 0:End Sub

'****************
'Gofer Animations
'****************

Dim LGPos, LGDir, RGPos, RGPos2, RGDir, LrampPos, LrampDir, RrampPos, RrampDir
LGPos=22:RGPos=22:LrampPos=22:RrampPos=22
LGDir=0:RGDir=0:LrampDir=1:RrampDir=1
LGofer.TimerEnabled=1:RGofer.TimerEnabled=1:LrampTimer.Enabled=1

    Sub LGoferUp_Timer()
        Select Case LGPos
          Case 0: LeftGoferBlock.z=-20
                LGoferIsUp = 1
                LGoferUp.Enabled = 0
                Lramp.HeightBottom=110
          Case 1: LeftGoferBlock.z=-25:Lramp.HeightBottom=104
          Case 2: LeftGoferBlock.z=-31:Lramp.HeightBottom=102
          Case 3: LeftGoferBlock.z=-37:Lramp.HeightBottom=100
          Case 4: LeftGoferBlock.z=-31:Lramp.HeightBottom=102
          Case 5: LeftGoferBlock.z=-25:Lramp.HeightBottom=104
          Case 6: LeftGoferBlock.z=-20:Lramp.HeightBottom=110
          Case 7: LeftGoferBlock.z=-25:Lramp.HeightBottom=104
          Case 8: LeftGoferBlock.z=-31:Lramp.HeightBottom=98
          Case 9: LeftGoferBlock.z=-37:Lramp.HeightBottom=91
          Case 10: LeftGoferBlock.z=-43:Lramp.HeightBottom=84
          Case 11: LeftGoferBlock.z=-49:Lramp.HeightBottom=77
          Case 12: LeftGoferBlock.z=-55:Lramp.HeightBottom=70
          Case 13: LeftGoferBlock.z=-61:Lramp.HeightBottom=63
          Case 14: LeftGoferBlock.z=-67:Lramp.HeightBottom=56
          Case 15: LeftGoferBlock.z=-73:Lramp.HeightBottom=49
          Case 16: LeftGoferBlock.z=-79:Lramp.HeightBottom=42
          Case 17: LeftGoferBlock.z=-85:Lramp.HeightBottom=35
          Case 18: LeftGoferBlock.z=-91:Lramp.HeightBottom=28
          Case 19: LeftGoferBlock.z=-97:Lramp.HeightBottom=21
          Case 20: LeftGoferBlock.z=-103:Lramp.HeightBottom=14
          Case 21: LeftGoferBlock.z=-109:Lramp.HeightBottom=7
          Case 22: LeftGoferBlock.z=-115:Lramp.HeightBottom=0
        End Select
          If LGpos>0 then LGpos=LGpos-1
    End Sub

    Sub LGoferDown_Timer()
    Select Case LGPos
      Case 0: LeftGoferBlock.z=-19
      Case 1: LeftGoferBlock.z=-25
      Case 2: LeftGoferBlock.z=-31
      Case 3: LeftGoferBlock.z=-37
      Case 4: LeftGoferBlock.z=-43
      Case 5: LeftGoferBlock.z=-49
      Case 6: LeftGoferBlock.z=-55
      Case 7: LeftGoferBlock.z=-61
      Case 8: LeftGoferBlock.z=-67
      Case 9: LeftGoferBlock.z=-73
      Case 10: LeftGoferBlock.z=-79
      Case 11: LeftGoferBlock.z=-85
      Case 12: LeftGoferBlock.z=-91
      Case 13: LeftGoferBlock.z=-97
      Case 14: LeftGoferBlock.z=-103
      Case 15: LeftGoferBlock.z=-109
      Case 16: LeftGoferBlock.z=-115
            LGoferDown.Enabled = 0
            LGoferIsUp = 0
    End Select
      If LGpos<16 then LGpos=LGpos+1
    End Sub

    Sub RGoferUp_Timer()
        Select Case RGPos
          Case 0: RightGoferBlock.z=-20
                RGoferIsUp = 1
                RGoferUp.Enabled = 0
                Rramp.HeightBottom=110
          Case 1: RightGoferBlock.z=-25:Rramp.HeightBottom=104
          Case 2: RightGoferBlock.z=-31:Rramp.HeightBottom=102
          Case 3: RightGoferBlock.z=-37:Rramp.HeightBottom=100
          Case 4: RightGoferBlock.z=-31:Rramp.HeightBottom=102
          Case 5: RightGoferBlock.z=-25:Rramp.HeightBottom=104
          Case 6: RightGoferBlock.z=-20:Rramp.HeightBottom=110
          Case 7: RightGoferBlock.z=-25:Rramp.HeightBottom=104
          Case 8: RightGoferBlock.z=-31:Rramp.HeightBottom=98
          Case 9: RightGoferBlock.z=-37:Rramp.HeightBottom=91
          Case 10: RightGoferBlock.z=-43:Rramp.HeightBottom=84
          Case 11: RightGoferBlock.z=-49:Rramp.HeightBottom=77
          Case 12: RightGoferBlock.z=-55:Rramp.HeightBottom=70
          Case 13: RightGoferBlock.z=-61:Rramp.HeightBottom=63
          Case 14: RightGoferBlock.z=-67:Rramp.HeightBottom=56
          Case 15: RightGoferBlock.z=-73:Rramp.HeightBottom=49
          Case 16: RightGoferBlock.z=-79:Rramp.HeightBottom=42
          Case 17: RightGoferBlock.z=-85:Rramp.HeightBottom=35
          Case 18: RightGoferBlock.z=-91:Rramp.HeightBottom=28
          Case 19: RightGoferBlock.z=-97:Rramp.HeightBottom=21
          Case 20: RightGoferBlock.z=-103:Rramp.HeightBottom=14
          Case 21: RightGoferBlock.z=-109:Rramp.HeightBottom=7
          Case 22: RightGoferBlock.z=-115:Rramp.HeightBottom=0
        End Select
          If RGpos>0 then RGpos=RGpos-1
    End Sub

    Sub RGoferDown_Timer()
    Select Case RGPos
      Case 0: RightGoferBlock.z=-19
      Case 1: RightGoferBlock.z=-25
      Case 2: RightGoferBlock.z=-31
      Case 3: RightGoferBlock.z=-37
      Case 4: RightGoferBlock.z=-43
      Case 5: RightGoferBlock.z=-49
      Case 6: RightGoferBlock.z=-55
      Case 7: RightGoferBlock.z=-61
      Case 8: RightGoferBlock.z=-67
      Case 9: RightGoferBlock.z=-73
      Case 10: RightGoferBlock.z=-79
      Case 11: RightGoferBlock.z=-85
      Case 12: RightGoferBlock.z=-91
      Case 13: RightGoferBlock.z=-97
      Case 14: RightGoferBlock.z=-103
      Case 15: RightGoferBlock.z=-109
      Case 16: RightGoferBlock.z=-115
            RGoferDown.Enabled = 0
            RGoferIsUp = 0
    End Select
      If RGpos<16 then RGpos=RGpos+1
    End Sub

Sub LrampTimer_Timer()
  If LGoferIsUp = 0 Then
    Select Case LrampPos
          Case 0: Lramp.HeightBottom=110
          Case 1: Lramp.HeightBottom=104
          Case 2: Lramp.HeightBottom=102
          Case 3: Lramp.HeightBottom=100
          Case 4: Lramp.HeightBottom=102
          Case 5: Lramp.HeightBottom=104
          Case 6: Lramp.HeightBottom=110
          Case 7: Lramp.HeightBottom=104
          Case 8: Lramp.HeightBottom=98
          Case 9: Lramp.HeightBottom=91
          Case 10: Lramp.HeightBottom=84
          Case 11: Lramp.HeightBottom=77
          Case 12: Lramp.HeightBottom=70
          Case 13: Lramp.HeightBottom=63
          Case 14: Lramp.HeightBottom=56
          Case 15: Lramp.HeightBottom=49
          Case 16: Lramp.HeightBottom=42
          Case 17: Lramp.HeightBottom=35
          Case 18: Lramp.HeightBottom=28
          Case 19: Lramp.HeightBottom=21
          Case 20: Lramp.HeightBottom=14
          Case 21: Lramp.HeightBottom=7
          Case 22: Lramp.HeightBottom=0:Controller.Switch(47)=1:LeftRampUp = 0:LrampTimer.Enabled = 0
      If Lramppos<22 then Lramppos=Lramppos+1
    End Select
    End If
    End Sub


Sub RrampTimer_Timer()
  If RGoferIsUp = 0 Then
    Select Case RrampPos
          Case 0: Rramp.HeightBottom=110
          Case 1: Rramp.HeightBottom=104
          Case 2: Rramp.HeightBottom=102
          Case 3: Rramp.HeightBottom=100
          Case 4: Rramp.HeightBottom=102
          Case 5: Rramp.HeightBottom=104
          Case 6: Rramp.HeightBottom=110
          Case 7: Rramp.HeightBottom=104
          Case 8: Rramp.HeightBottom=98
          Case 9: Rramp.HeightBottom=91
          Case 10: Rramp.HeightBottom=84
          Case 11: Rramp.HeightBottom=77
          Case 12: Rramp.HeightBottom=70
          Case 13: Rramp.HeightBottom=63
          Case 14: Rramp.HeightBottom=56
          Case 15: Rramp.HeightBottom=49
          Case 16: Rramp.HeightBottom=42
          Case 17: Rramp.HeightBottom=35
          Case 18: Rramp.HeightBottom=28
          Case 19: Rramp.HeightBottom=21
          Case 20: Rramp.HeightBottom=14
          Case 21: Rramp.HeightBottom=7
          Case 22: Rramp.HeightBottom=0:Controller.Switch(48)=1:RightRampUp = 0:RrampTimer.Enabled = 0
      If Rramppos<22 then Rramppos=Rramppos+1
    End Select
    End If
    End Sub

'**************
'Turntable Subs
'**************

    Set MotorCallback=GetRef("SRPRoutine")
    Dim WheelNewPos,WheelOldPos
    WheelOldPos=0

    Sub SRPRoutine
      WheelNewPos=ABS(INT(Controller.GetMech(0)/4))
      If WheelNewPos<>WheelOldPos Then
        'disc.roty = DiskArray(WheelNewPos)
        If WheelNewPos>WheelOldPos And WheelNewPos<15 Then:mTT.SolMotorState False,True:TTtimer.Enabled=0:BlurDir = 0:Blur.Enabled = 1:End If
        If WheelNewPos<WheelOldPos And WheelNewPos>0 Then:mTT.SolMotorState True,True:TTtimer.Enabled=0:BlurDir = 0:Blur.Enabled = 1:End If
      End If

      disc.roty = 360 - (Controller.GetMech(0) /64 * 360) + 12

      WheelOldPos=WheelNewPos
    End Sub
'   Sub SRPRoutine
'     WheelNewPos=ABS(INT(Controller.GetMech(0)/4))
'     If WheelNewPos<>WheelOldPos Then
'       disc.roty = DiskArray(WheelNewPos)
'       If WheelNewPos>WheelOldPos And WheelNewPos<15 Then:mTT.SolMotorState False,True:TTtimer.Enabled=0:BlurDir = 0:Blur.Enabled = 1:End If
'       If WheelNewPos<WheelOldPos And WheelNewPos>0 Then:mTT.SolMotorState True,True:TTtimer.Enabled=0:BlurDir = 0:Blur.Enabled = 1:End If
'     End If
'     WheelOldPos=WheelNewPos
'   End Sub

    Sub TTtimer_Timer
      mTT.SolMotorState False,False
      mTT.SolMotorState True,False
      BlurDir = 1
      Blur.Enabled = 1
      TTtimer.Enabled=0
    End Sub
  Dim DiskArray, BlurDir, Blurlevel
    DiskArray=Array("12","350","335","305","282","258","235","215","190","168","146","125","100","80","55","32")
    Blurlevel = 0
  Sub Blur_Timer()
    Select Case Blurlevel
      Case 0: disc.image = "NGG_Map"
        If BlurDir=1 Then
          Blur.Enabled = 0
          TTtimer.Enabled=1
        End If
      Case 1: disc.image = "NGG_Map_Mblur1"
      Case 2: disc.image = "NGG_Map_Mblur2"
      Case 3: disc.image = "NGG_Map_Mblur3"
      Case 4: disc.image = "NGG_Map_Mblur4"
      Case 5: disc.image = "NGG_Map_Mblur5"
      Case 6: disc.image = "NGG_Map_Mblur6"
        If BlurDir = 0 Then
          Blur.Enabled = 0
          TTtimer.Enabled=1
        End If
    End Select
    If BlurDir = 1 Then
      If Blurlevel>0 then Blurlevel = Blurlevel - 1
    Else
      If Blurlevel<6 Then Blurlevel = Blurlevel + 1
    End If
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
LampTimer.Interval = 10 'lamp fading speed
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


Sub UpdateLamps
  NFadeL 11, L11
  NFadeL 12, L12
  NFadeL 13, L13
  NFadeL 14, L14
  NFadeL 15, L15
  NFadeL 16, L16
  NFadeL 17, L17
  NFadeL 18, L18
  NFadeL 21, L21
  NFadeL 22, L22
  NFadeL 23, L23
  NFadeL 24, L24
  NFadeL 25, L25
  NFadeL 26, L26
  NFadeL 27, L27
  NFadeL 28, L28
  NFadeL 31, L31
  NFadeL 32, L32
  NFadeL 33, L33
  NFadeL 34, L34
  NFadeL 35, L35
  NFadeL 36, L36
  NFadeL 37, L37
  NFadeL 38, L38
  NFadeL 41, L41
  NFadeL 42, L42
  NFadeL 43, L43
  NFadeL 44, L44
  NFadeL 45, L45
  NFadeL 46, L46
  NFadeL 47, L47
  NFadeL 48, L48
  NFadeL 51, L51
  NFadeL 52, L52
  NFadeL 53, L53
  NFadeL 54, L54
  NFadeLm 55, L55b
  NFadeL 55, L55
  NFadeL 56, L56
  NFadeL 57, L57
  NFadeL 58, L58
  NFadeL 61, L61
  NFadeL 62, L62
  NFadeL 63, L63
  NFadeL 64, L64
  NFadeL 65, L65
  NFadeL 66, L66
  NFadeL 67, L67
  NFadeL 68, L68
  NFadeLm 71, L71b
  NFadeL 71, L71
  NFadeL 72, L72
  NFadeL 73, L73
  NFadeL 74, L74
  NFadeL 75, L75
  NFadeL 76, L76
  NFadeL 77, L77
  NFadeLm 78, bump3light1
  NFadeL 78, bump3light
  NFadeL 81, L81
  NFadeL 82, L82
  NFadeL 83, L83
  NFadeL 84, L84
  NFadeL 85, L85
  NFadeLm 86, bump2light1
  NFadeL 86, bump2light
  NFadeLm 87, bump1light1
  NFadeL 87, bump1light
  NFadeL 88, L88  'start button
  NFadeL 122, L122
  NFadeLm 131, L131b
  NFadeLm 131, L131a
  NFadeL 131, L131
  'NFadeL 145, L145
  'NFadeL 146, L146
  NFadeARmCombi 145, 146, upperplayfield, "upperplayfield", "upperplayfieldright", "upperplayfieldleft", "upperplayfieldboth", Combifor145and146
  NFadeLm 170, L170b
  NFadeLm 170, L170a
  NFadeL 170, L170
  NFadeLm 172, L172a
  NFadeLm 172, L172b
  NFadeL 172, L172
  NFadeLm 173, L173a
  NFadeLm 173, L173b
  NFadeL 173, L173
  NFadeLm 174, L174a
  NFadeLm 174, L174b
  NFadeL 174, L174
  NFadeLm 175, L175a
  NFadeLm 175, L175b
  NFadeL 175, L175
  NFadeLm 176, L176a
  NFadeLm 176, L176b
  NFadeL 176, L176
  NFadeLm 177, L177a
  NFadeLm 177, L177b
  NFadeL 177, L177
  NFadeLm 178, L178a
  NFadeLm 178, L178b
  NFadeL 178, L178
  NFadeLm 179, L179a
  NFadeLm 179, L179b
  NFadeL 179, L179
  NFadeLm 180, L180a
  NFadeLm 180, L180b
  NFadeLm 180, L180c
  NFadeL 180, L180
  NFadeLm 190, L190
  NFadeLm 190, L190a
  NFadeL 190, L190b
End Sub
'
'Sub FlasherTimer_Timer()
'   Flash 11, h11
'   Flash 12, h12
'   Flash 13, h13
'   Flash 14, h14
'   Flash 15, h15
'   Flash 16, h16
'   Flash 17, h17
'   Flash 18, h18
'   Flash 21, h21
'   Flash 22, h22
'   Flash 23, h23
'   Flash 24, h24
'   Flash 25, h25
'   Flash 26, h26
'   Flash 27, h27
'   Flash 28, h28
'   Flash 31, h31
'   Flash 32, h32
'   Flash 33, h33
'   Flash 34, h34
'   Flash 35, h35
'   Flash 36, h36
'   Flash 37, h37
'   Flash 38, h38
'   Flash 41, h41
'   Flash 42, h42
'   Flash 43, h43
'   Flash 44, h44
'   Flash 45, h45
'   Flash 46, h46
'   Flash 47, h47
'   Flash 48, h48
'   Flash 51, h51
'   Flash 52, h52
'   Flash 53, h53
'   Flash 54, h54
'   Flashm 55, h55a
'   Flash 55, h55
'   Flash 56, h56
'   Flash 57, h57
'   Flash 58, h58
'   Flash 61, h61
'   Flash 62, h62
'   Flash 63, h63
'   Flash 64, h64
'   Flash 65, h65
'   Flash 66, h66
'   Flash 68, h68
'   Flashm 71, h71a
'   Flash 71, h71
'   Flash 72, h72
'   Flash 73, h73
'   Flash 74, h74
'   Flash 75, h75
'   Flash 77, h77
'   Flash 78, h78
'   Flash 81, h81
'   Flash 82, h82
'   Flash 83, h83
'   Flash 84, h84
'   Flash 85, h85
'   Flash 86, h86
'   Flash 87, h87
'   Flashm 120, Flasher19
'   Flash 120, Flasher1
'   Flashm 121, Flasher20
'   Flash 121, Flasher2
'   Flashm 122, Flasher21
'   Flashm 122, Flasher4
'   Flash 122, Flasher3
'   Flash 123, Flasher5
'   Flashm 124, Flasher15
'   Flash 124, Flasher6
'   Flashm 125, Flasher18
'   Flash 125, Flasher7
'   Flash 126, Flasher8
'   Flash 127, Flasher9
'   Flashm 128, Flasher17
'   Flash 128, Flasher10
'   Flashm 129, Flasher16
'   Flash 129, Flasher11
'   Flash 130, Flasher12
'   Flashm 131, Flasher14
'   Flash 131, Flasher13
' End Sub
'

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
        Case 6, 7.8:FadingLevel(nr) = FadingLevel(nr) + 1             'wait
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

' Desktop Objects: Reels & texts (you may also use lights on the desktop)

' Reels
Dim NewCombiResult, Combifor145and146
Sub NFadeARmCombi(nr1, nr2, ramp, aOff, a1, a2, aBoth, OldCombiResult)
  NewCombiResult=LampState(nr1)+(2*LampState(nr2))
  If OldCombiResult=NewCombiResult Then Exit Sub
  Select Case NewCombiResult
    Case 0:ramp.image=aOff
    Case 1:ramp.image=a1
    Case 2:ramp.image=a2
    Case 3:ramp.image=aBoth
  End Select
  OldCombiResult=NewCombiResult
End Sub

Dim LockbarEnabled

Sub FadeR(nr, object)
    Select Case FadingLevel(nr)
        Case 4:object.SetValue 1:FadingLevel(nr) = 6                   'fading to off...
        Case 5:object.SetValue 0:FadingLevel(nr) = 1                   'ON
        Case 6, 7.8:FadingLevel(nr) = FadingLevel(nr) + 1              'wait
        Case 9:object.SetValue 2:FadingLevel(nr) = FadingLevel(nr) + 1 'fading...
        Case 10, 11, 12:FadingLevel(nr) = FadingLevel(nr) + 1          'wait
        Case 13:object.SetValue 3:FadingLevel(nr) = 0                  'Off
    End Select
End Sub

Sub FadeRm(nr, object)
    Select Case FadingLevel(nr)
        Case 4:object.SetValue 1
        Case 5:object.SetValue 0
        Case 9:object.SetValue 2
        Case 3:object.SetValue 3
    End Select
End Sub

'Texts

Sub NFadeT(nr, object, message)
    Select Case FadingLevel(nr)
        Case 4:object.Text = "":FadingLevel(nr) = 0
        Case 5:object.Text = message:FadingLevel(nr) = 1
    End Select
End Sub

Sub NFadeTm(nr, object, b)
    Select Case FadingLevel(nr)
        Case 4:object.Text = ""
        Case 5:object.Text = message
    End Select
End Sub

'***********
' Update GI
'***********
Dim gistep, Giswitch
Giswitch = 0

Sub UpdateGIon(no, Enabled)
  Select Case no
    Case 0 'GIright
      If Enabled Then
        For each xx in GIright:xx.State = 1: Next
        TargetBrightness = 1.0 ' Full brightness when GI ON
        If VRRoom > 0 Then
          VRBGLit.visible = 1
        End If
      Else
        For each xx in GIright:xx.state = 0: Next
        TargetBrightness = 0.4 ' Dimmed when GI OFF
        If VRRoom > 0 Then
          VRBGLit.visible = 0
        End If
        End If

      ' Start the brightness fade timer
      If Not BrightnessTimer.Enabled Then BrightnessTimer.Enabled = True

      ' Track GIright state
      GIrightEnabled = Enabled

    Case 1 'GIleft
      If Enabled Then
        For Each xx in GIleft:xx.state = 1: Next
      else
        For Each xx in GIleft:xx.state = 0: Next
      End If

      ' Track GIleft state
      GIleftEnabled = Enabled

    Case 2 'GIother
      If Enabled Then
        For Each xx in GIother:xx.state = 1: Next
      Else
        For Each xx in GIother:xx.state = 0: Next
      End If
  End Select

  ' Update blade visibility whenever GIright OR GIleft changes
  If no = 0 OR no = 1 Then
    UpdateBladeVisibility
  End If
End Sub

Sub UpdateBladeVisibility()
    ' First check if blades should be visible at all
    If VRRoom = 0 And Table.ShowDT = False And VRCab_BladesEnabled = 0 Then
        VRCab_Blades.visible = 0
        VRCab_Blades_OFF.visible = 0
        Exit Sub
    End If

    ' Show correct blades based on GI state AND selected style

    If GIrightEnabled OR GIleftEnabled Then
        ' EITHER GIright OR GIleft is ON - show VRCab_Blades
        VRCab_Blades.visible = 1
        VRCab_Blades_OFF.visible = 0

        ' Apply correct image based on style selection
        If sidebladeStyle = 1 Then
            ' Normal style
            VRCab_Blades.Image = "VRCab_Blades_Normal"
        Else
            ' Artwork style
            VRCab_Blades.Image = "VRCab_Blades_Artwork"
        End If

    Else
        ' BOTH GIright AND GIleft are OFF - show VRCab_Blades_OFF
        VRCab_Blades.visible = 0
        VRCab_Blades_OFF.visible = 1

        ' Apply correct image based on style selection
        If sidebladeStyle = 1 Then
            ' Normal style
            VRCab_Blades_OFF.Image = "VRCab_Blades_OFF_Normal"
        Else
            ' Artwork style
            VRCab_Blades_OFF.Image = "VRCab_Blades_OFF_Artwork"
        End If
    End If
End Sub

Sub UpdateApronWall()
    ' Update apron1 wall texture based on style selection

    ' apron1 should already exist as a wall object
    If Not apron1 Is Nothing Then
        If apronWallStyle = 1 Then
            ' Style 1 texture
            apron1.Image = "apron_style1" ' Your image name for style 1
        Else
            ' Style 2 texture
            apron1.Image = "apron_style2" ' Your image name for style 2
        End If
    End If

    ' Optional debug
    ' Debug.Print "Apron wall set to style: " & apronWallStyle
End Sub



Sub UpdateGI(no, step)
    If step = 0 then exit sub 'only values from 1 to 8 are visible and reliable. 0 is not reliable and 7 & 8 are the same so...

    gistep = (step-1) / 7

    ' When GI intensity changes, update the blade visibility
    ' Track which GI bank is being adjusted
    If no = 0 Then ' GIright intensity
        If step > 0 Then
            GIrightEnabled = True
        Else
            GIrightEnabled = False
        End If
        UpdateBladeVisibility
    ElseIf no = 1 Then ' GIleft intensity
        If step > 0 Then
            GIleftEnabled = True
        Else
            GIleftEnabled = False
        End If
        UpdateBladeVisibility
    End If

    ' Rest of your existing UpdateGI code...
    Dim xx ' Use xx instead of ii to match your other code
    Select Case no
        Case 0
          For each xx in GIleft
                xx.IntensityScale = gistep
            Next
        Case 1
            For each xx in GIright
                xx.IntensityScale = gistep
            Next
        Case 2 ' also the bumpers er GI
            For each xx in GIother
                xx.IntensityScale = gistep
            Next
            Select case gistep
                Case 0
                    GolfCart.image = "GolfCartMap"
                Case 1
                    GolfCart.image = "GolfCartMap_ON1"
                Case 2
                    GolfCart.image = "GolfCartMap_ON1"
                Case 3
                    GolfCart.image = "GolfCartMap_ON2"
                Case 4
                    GolfCart.image = "GolfCartMap_ON3"
                Case 5
                    GolfCart.image = "GolfCartMap_ON4"
                Case 6
                    GolfCart.image = "GolfCartMap_ON4"
                Case 7
                    GolfCart.image = "GolfCartMap_ON5"
            End Select
    End Select
'    ' change the intensity of the flasher depending on the gi to compensate for the gi lights being off
'    For ii = 0 to 200
'        FlashMax(ii) = 6 - gistep * 3 ' the maximum value of the flashers
'    Next
End Sub


'''Ramp Helpers

'Sub centerramphelper_Hit():ActiveBall.Vely=ActiveBall.Vely*1.3:End Sub
'Sub centerramphelper2_Hit():Activeball.Velx=ActiveBall.Velx-20:End Sub
'Sub upperhelper_Hit():ActiveBall.Velx=ActiveBall.Velx*1.2:End Sub
'Sub upperhelper2_Hit():ActiveBall.Vely=ActiveBall.Vely*1.3:End Sub
'Sub jumpramptrigger_Hit():ActiveBall.Vely=ActiveBall.Vely*1.3:End Sub
'Sub upperjump_Hit():ActiveBall.Vely=ActiveBall.Vely*1.5:End Sub
'Sub upperjump_UnHit():ActiveBall.Velz=ActiveBall.Velz=0:End Sub


Sub Game_timer()
  slamrampP.z = Flipper1.CurrentAngle
  If RightRampUp = 0 Then RrampTimer.Enabled = 1
  If LeftRampUp = 0 Then LRampTimer.Enabled = 1
End Sub


'*****************************************
'      JP's VP10 Rolling Sounds
'*****************************************

'Const tnob = 8 ' total number of balls
'ReDim rolling(tnob)
'InitRolling
'
'Sub InitRolling
'    Dim i
'    For i = 0 to tnob
'        rolling(i) = False
'    Next
'End Sub
'
'Sub RollingTimer_Timer()
'    Dim BOT, b
'    BOT = GetBalls
'
' ' stop the sound of deleted balls
'    For b = UBound(BOT) + 1 to tnob
'        rolling(b) = False
'        StopSound("fx_ballrolling" & b)
'    Next

' ' exit the sub if no balls on the table
'    If UBound(BOT) = -1 Then Exit Sub

  ' play the rolling sound for each ball

'    For b = 0 to UBound(BOT)
'      If BallVel(BOT(b) ) > 1 Then
'        rolling(b) = True
'        if BOT(b).z < 30 Then ' Ball on playfield
'          PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) ), AudioPan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0, AudioFade(BOT(b) )
'        Else ' Ball on raised ramp
'          PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) )*.5, AudioPan(BOT(b) ), 0, Pitch(BOT(b) )+50000, 1, 0, AudioFade(BOT(b) )
'        End If
'      Else
'        If rolling(b) = True Then
'          StopSound("fx_ballrolling" & b)
'          rolling(b) = False
'        End If
'      End If
'    Next
'End Sub

'**********************
' Ball Collision Sound
'**********************

'Sub OnBallBallCollision(ball1, ball2, velocity)
'    FlipperCradleCollision ball1, ball2, velocity
' PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 2000, AudioPan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
'End Sub

 '**********************
'Flipper Shadows
'***********************
Sub RealTime_Timer
  lfs.RotZ = LeftFlipper.CurrentAngle
  rfs.RotZ = RightFlipper.CurrentAngle
BallShadowUpdate
End Sub


Sub BallShadowUpdate()
Dim BallShadow
BallShadow = Array (BallShadow1,BallShadow2,BallShadow3,BallShadow4,BallShadow5,BallShadow6,BallShadow7,BallShadow8)
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
    BallShadow(b).X = BOT(b).X
    ballShadow(b).Y = BOT(b).Y + 10
        If BOT(b).Z > 20 and BOT(b).Z < 200 Then
            BallShadow(b).visible = 1
        Else
            BallShadow(b).visible = 0
        End If
if BOT(b).z > 30 Then
ballShadow(b).height = BOT(b).Z - 20
ballShadow(b).opacity = 110
Else
ballShadow(b).height = BOT(b).Z - 20
ballShadow(b).opacity = 90
End If
    Next
End Sub

Sub platformdrop_Unhit()
  If ActiveBall.Vely>0 then PlaySound "ball_bounce"
End Sub

Sub rightrampdrop_UnHit():PlaySound "ball_bounce":End Sub
Sub leftrampdrop_UnHit():PlaySound "ball_bounce":End Sub

Sub jumpramptrigger_Hit():PlaySound "metalhit_medium":End Sub






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
' - Nudging, plunger, coin-in, start button sounds will be added to the keydown and keyup subs.
' - Flipper sounds in the flipper solenoid subs. Flipper collision sounds in the flipper collide subs.
' - Bumpers, slingshots, drain, ball release, knocker, spinner, and saucers in their respective subs
' - Ball rolling sounds sub
'
' Tutorial videos by Apophis
' Audio : Adding Fleep Part 1         https://youtu.be/rG35JVHxtx4?si=zdN9W4cZWEyXbOz_
' Audio : Adding Fleep Part 2         https://youtu.be/dk110pWMxGo?si=2iGMImXXZ0SFKVCh
' Audio : Adding Fleep Part 3         https://youtu.be/ESXWGJZY_EI?si=6D20E2nUM-xAw7xy


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
SpinnerSoundLevel = 0.5      'volume level; range [0, 1]

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
  PlaySound playsoundparams, 0, min(aVol,1) * VolumeDial, AudioPan(tableobj), 0, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelExistingStatic(playsoundparams, aVol, tableobj)
  PlaySound playsoundparams, 0, min(aVol,1) * VolumeDial, AudioPan(tableobj), 0, 0, 1, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelStaticLoop(playsoundparams, aVol, tableobj)
  PlaySound playsoundparams, - 1, min(aVol,1) * VolumeDial, AudioPan(tableobj), 0, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelStaticRandomPitch(playsoundparams, aVol, randomPitch, tableobj)
  PlaySound playsoundparams, 0, min(aVol,1) * VolumeDial, AudioPan(tableobj), randomPitch, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelActiveBall(playsoundparams, aVol)
  PlaySound playsoundparams, 0, min(aVol,1) * VolumeDial, AudioPan(ActiveBall), 0, 0, 0, 0, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtLevelExistingActiveBall(playsoundparams, aVol)
  PlaySound playsoundparams, 0, min(aVol,1) * VolumeDial, AudioPan(ActiveBall), 0, 0, 1, 0, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtLeveTimerActiveBall(playsoundparams, aVol, ballvariable)
  PlaySound playsoundparams, 0, min(aVol,1) * VolumeDial, AudioPan(ballvariable), 0, 0, 0, 0, AudioFade(ballvariable)
End Sub

Sub PlaySoundAtLevelTimerExistingActiveBall(playsoundparams, aVol, ballvariable)
  PlaySound playsoundparams, 0, min(aVol,1) * VolumeDial, AudioPan(ballvariable), 0, 0, 1, 0, AudioFade(ballvariable)
End Sub

Sub PlaySoundAtLevelRoll(playsoundparams, aVol, pitch)
  PlaySound playsoundparams, - 1, min(aVol,1) * VolumeDial, AudioPan(tableobj), randomPitch, 0, 0, 0, AudioFade(tableobj)
End Sub

' Previous Positional Sound Subs

Sub PlaySoundAt(soundname, tableobj)
  PlaySound soundname, 1, 1 * VolumeDial, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

Sub PlaySoundAtVol(soundname, tableobj, aVol)
  PlaySound soundname, 1, min(aVol,1) * VolumeDial, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

Sub PlaySoundAtBall(soundname)
  PlaySoundAt soundname, ActiveBall
End Sub

Sub PlaySoundAtBallVol (Soundname, aVol)
  PlaySound soundname, 1,min(aVol,1) * VolumeDial, AudioPan(ActiveBall), 0,0,0, 1, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtBallVolM (Soundname, aVol)
  PlaySound soundname, 1,min(aVol,1) * VolumeDial, AudioPan(ActiveBall), 0,0,0, 0, AudioFade(ActiveBall)
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

Sub Posts_Hit(idx)
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

'/////////////////////////////  GI AND FLASHER   ////////////////////////////

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
'****  END FLEEP MECHANICAL SOUNDS
'******************************************************

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

'*******************************************
'  ZOPT: User Options
'*******************************************

Dim LightLevel : LightLevel = 0.25        ' Level of room lighting (0 to 1), where 0 is dark and 100 is brightest
Dim ColorLUT : ColorLUT = 5           ' Color desaturation LUTs: 1 to 11, where 1 is normal and 11 is black'n'white
Dim VolumeDial : VolumeDial = 0.8             ' Overall Mechanical sound effect volume. Recommended values should be no greater than 1.
Dim BallRollVolume : BallRollVolume = 0.5     ' Level of ball rolling volume. Value between 0 and 1
Dim RampRollVolume : RampRollVolume = 0.5     ' Level of ramp rolling volume. Value between 0 and 1
'Dim StagedFlippers                       ' Staged Flippers. 0 = Disabled, 1 = Enabled
Dim VRTest : VRTest = False

          ' Display VR in Desktop Mode
Dim VRRoom
Dim VRRoomChoice : VRRoomChoice = 1 ' 1 - Minimal Room, 2 - Ultra Minimal, 3 - MEGA
Dim BallType
Dim LastBallType
Dim TargetBrightness, CurrentBrightness
Dim BrightnessFadeSpeed

' Called when options are tweaked by the player.
' - 0: game has started, good time to load options and adjust accordingly
' - 1: an option has changed
' - 2: options have been reseted
' - 3: player closed the tweak UI, good time to update staticly prerendered parts
' Table1.Option arguments are:
' - option name, minimum value, maximum value, step between valid values, default value, unit (0=None, 1=Percent), an optional arry of literal strings

Dim dspTriggered : dspTriggered = False
Sub Table_OptionEvent(ByVal eventId)
    If eventId = 1 And Not dspTriggered Then dspTriggered = True : DisableStaticPreRendering = True : End If

  Dim Anaglyph
    Anaglyph = 0 ' Default to 0 (no anaglyph)
     ' Color Saturation
    ColorLUT = Table.Option("Color Saturation", 1, 27, 1, 1, 0, _
        Array("Normal", "Desaturated 10%", "Desaturated 20%", "Desaturated 30%", "Desaturated 40%", "Desaturated 50%", "Desaturated 60%", "Desaturated 70%", "Desaturated 80%", "Desaturated 90%", "Black 'n White", _
        "Fleep Natural Dark 1", "Fleep Natural Dark 2", "Fleep Warm Dark", "Fleep Warm Bright", "Fleep Warm Vivid Soft", "Fleep Warm Vivid Hard", "Skitso Natural & Balance", "Skitso Natural High Contrast", _
        "3rdaxis THX Standard", "Callev Brightness & Contrast", "Hauntfreaks Desaturated", "Tomate Washed out", "VPW Original 1 on 1", "Bassgeige", "Blacklight", "B&W Comic Book"))

    if ColorLUT = 1 Then Table.ColorGradeImage = "" ' "Normal"
    if ColorLUT = 2 Then Table.ColorGradeImage = "colorgradelut256x16-10" ' "Desaturated 10%"
    if ColorLUT = 3 Then Table.ColorGradeImage = "colorgradelut256x16-20" ' "Desaturated 20%"
    if ColorLUT = 4 Then Table.ColorGradeImage = "colorgradelut256x16-30" ' "Desaturated 30%"
    if ColorLUT = 5 Then Table.ColorGradeImage = "colorgradelut256x16-40" ' "Desaturated 40%"
    if ColorLUT = 6 Then Table.ColorGradeImage = "colorgradelut256x16-50" ' "Desaturated 50%"
    if ColorLUT = 7 Then Table.ColorGradeImage = "colorgradelut256x16-60" ' "Desaturated 60%"
    if ColorLUT = 8 Then Table.ColorGradeImage = "colorgradelut256x16-70" ' "Desaturated 70%"
    if ColorLUT = 9 Then Table.ColorGradeImage = "colorgradelut256x16-80" ' "Desaturated 80%"
    if ColorLUT = 10 Then Table.ColorGradeImage = "colorgradelut256x16-90" ' "Desaturated 90%"
    if ColorLUT = 11 Then Table.ColorGradeImage = "colorgradelut256x16-100" ' "Black 'n White"
    if ColorLUT = 12 Then Table.ColorGradeImage = "colorgradelut256x16-fleep-natural-dark-1" ' "Fleep Natural Dark 1"
    if ColorLUT = 13 Then Table.ColorGradeImage = "colorgradelut256x16-fleep-natural-dark-2" ' "Fleep Natural Dark 2"
    if ColorLUT = 14 Then Table.ColorGradeImage = "colorgradelut256x16-fleep-warm-dark" ' "Fleep Warm Dark "
    if ColorLUT = 15 Then Table.ColorGradeImage = "colorgradelut256x16-fleep-warm-bright" ' "Fleep Warm Bright"
    if ColorLUT = 16 Then Table.ColorGradeImage = "colorgradelut256x16-fleep-warm-vivid-soft" ' "Fleep Warm Vivid Soft"
    if ColorLUT = 17 Then Table.ColorGradeImage = "colorgradelut256x16-fleep-warm-vivid-hard" ' "Fleep Warm Vivid Hard"
    if ColorLUT = 18 Then Table.ColorGradeImage = "colorgradelut256x16-skitso-natural-and-balance" ' "Skitso Natural & Balance"
    if ColorLUT = 19 Then Table.ColorGradeImage = "colorgradelut256x16-skitso-natural-high-contrast" ' "Skitso Natural High Contrast"
    if ColorLUT = 20 Then Table.ColorGradeImage = "colorgradelut256x16-3rdaxis-thx-standard" ' "3rdaxis THX Standard"
    if ColorLUT = 21 Then Table.ColorGradeImage = "colorgradelut256x16-callev-brightness-and-contrast" ' "Callev Brightness & Contrast"
    if ColorLUT = 22 Then Table.ColorGradeImage = "colorgradelut256x16-hauntfreaks-desaturated" ' "Hauntfreaks Desaturated"
    if ColorLUT = 23 Then Table.ColorGradeImage = "colorgradelut256x16-tomate-washed-out" ' "Tomate Washed out"
    if ColorLUT = 24 Then Table.ColorGradeImage = "colorgradelut256x16-vpw-original-1-on-1" ' "VPW Original 1 on 1"
    if ColorLUT = 25 Then Table.ColorGradeImage = "colorgradelut256x16-bassgeige" ' "Bassgeige"
    if ColorLUT = 26 Then Table.ColorGradeImage = "colorgradelut256x16-blacklight" ' "Blacklight"
    if ColorLUT = 27 Then Table.ColorGradeImage = "colorgradelut256x16-b-and-w-comic-book" ' "B&W Comic Book"



    BallType = Table.Option("Ball Type", 1, 2, 1, 1, 0, Array("Golf Ball", "Chrome Ball"))
    ApplyBallProperties

    apronWallStyle = Table.Option("Apron", 1, 2, 1, 1, 0, Array("Custom", "Original"))
    UpdateApronWall

    BackTopEnabled = Table.Option("Back Top", 0, 1, 1, 1, 0, Array("OFF", "ON"))

    LockbarEnabled = Table.Option("Lockbar/Rails", 0, 1, 1, 1, 0, Array("OFF", "ON"))

    sidebladeStyle = Table.Option("Side Blades", 1, 2, 1, 1, 0, Array("Normal", "Artwork"))
     UpdateBladeVisibility

     VRCab_BladesEnabled = Table.Option("Cabinet Blades", 0, 1, 1, 1, 0, Array("OFF", "ON"))

       SideFlashersEnabled = Table.Option("Side Flashers", 0, 1, 1, 1, 0, Array("OFF", "ON"))
       UpdateSideFlashers


  ' Room brightness
' LightLevel = Table1.Option("Table Brightness (Ambient Light Level)", 0, 1, 0.01, .5, 1)
  LightLevel = NightDay/100
' SetRoomBrightness LightLevel   'Uncomment this line for lightmapped tables.

    ' Staged Flippers
'    StagedFlippers = Table.Option("Staged Flippers", 0, 1, 1, 0, 0, Array("Disabled", "Enabled"))

    ' VRRoom
  VRRoomChoice = Table.Option("VR Room", 1, 3, 1, 3, 0, Array("Minimal", "Cab Only", "Golf Course"))

  LoadVRRoom
    UpdateBackTopVisibility
    ' VR Room
'    VRRoom = Table.Option("VR Room", 0, 2, 1, 2, 0, Array("Minimal", "Cab Only", "Golf Course"))
    If eventId = 3 And dspTriggered Then dspTriggered = False : DisableStaticPreRendering = False : End If

' Sound volumes
    VolumeDial = Table.Option("Mech Volume", 0, 1, 0.01, 0.8, 1)
    BallRollVolume = Table.Option("Ball Roll Volume", 0, 1, 0.01, 0.5, 1)
  RampRollVolume = Table.Option("Ramp Roll Volume", 0, 1, 0.01, 0.5, 1)

End Sub

'******************************************************
' ZBRL: BALL ROLLING AND DROP SOUNDS
'******************************************************
'
' Be sure to call RollingUpdate in a timer with a 10ms interval see the GameTimer_Timer() sub

ReDim rolling(tnob)
InitRolling

Dim DropCount
ReDim DropCount(tnob)

Sub InitRolling
  Dim i
  For i = 0 to tnob
    rolling(i) = False
  Next
End Sub

Sub RollingUpdate()
  Dim b, BOT
  BOT = GetBalls

  ' stop the sound of deleted balls
  For b = UBound(BOT) + 1 to tnob - 1
    rolling(b) = False
    StopSound("BallRoll_" & b)
  Next

  ' exit the sub if no balls on the table
  If UBound(BOT) = -1 Then Exit Sub

  ' play the rolling sound for each ball

  For b = 0 to UBound(BOT)
    If BallVel(BOT(b)) > 1 AND BOT(b).z < 30 Then
      rolling(b) = True
      PlaySound ("BallRoll_" & b), -1, VolPlayfieldRoll(BOT(b)) * BallRollVolume * VolumeDial, AudioPan(BOT(b)), 0, PitchPlayfieldRoll(BOT(b)), 1, 0, AudioFade(BOT(b))

    Else
      If rolling(b) = True Then
        StopSound("BallRoll_" & b)
        rolling(b) = False
      End If
    End If

    ' Ball Drop Sounds
    If BOT(b).VelZ < -1 and BOT(b).z < 55 and BOT(b).z > 27 Then 'height adjust for ball drop sounds
      If DropCount(b) >= 5 Then
        DropCount(b) = 0
        If BOT(b).velz > -7 Then
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
sub GameTimer_Timer()
  Cor.Update
  RollingUpdate

  ' Check if ball type changed and apply
  If LastBallType <> BallType Then
    ApplyBallProperties
    LastBallType = BallType
  End If
End Sub

Sub BrightnessTimer_Timer()
  ' Smoothly fade toward target brightness
  If CurrentBrightness < TargetBrightness Then
    CurrentBrightness = CurrentBrightness + BrightnessFadeSpeed
    If CurrentBrightness > TargetBrightness Then
      CurrentBrightness = TargetBrightness
    End If
  ElseIf CurrentBrightness > TargetBrightness Then
    CurrentBrightness = CurrentBrightness - BrightnessFadeSpeed
    If CurrentBrightness < TargetBrightness Then
      CurrentBrightness = TargetBrightness
    End If
  Else
    ' Reached target, stop the timer
    BrightnessTimer.Enabled = False
    Exit Sub
  End If

  ' Adjust ambient light intensity based on brightness
  Dim xx
  For Each xx in extraambient
    If IsObject(xx) Then
      xx.IntensityScale = CurrentBrightness
    End If
  Next

  ' Note: We don't adjust GI lights here because they're being turned on/off
  ' by the UpdateGIOn sub. The ambient lights provide the overall brightness.
End Sub

'******************************************************
'   ZRRL: RAMP ROLLING SFX
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


Function BallPitch(ball) ' Calculates the pitch of the sound based on the ball speed
    BallPitch = pSlope(BallVel(ball), 1, -1000, 60, 10000)
End Function

Function BallPitchV(ball) ' Calculates the pitch of the sound based on the ball speed Variation
  BallPitchV = pSlope(BallVel(ball), 1, -4000, 60, 7000)
End Function


'Ramp triggers


Sub RHP1_Hit()
    WireRampOn True
End Sub

Sub RHP010_Hit()
    WireRampOn True
End Sub


Sub RHP2_Hit
  WireRampOff
End Sub

Sub RHP001_Hit()
    WireRampOn True
End Sub

Sub RHP002_Hit
  WireRampOff
End Sub
Sub RHP00002_Hit
  If activeball.vely < 0 Then
  WireRampOff
end If
End Sub
Sub RHP000002_Hit
  If activeball.vely < 0 Then
  WireRampOff
end If
End Sub

Sub RHP0001_Hit()
    WireRampOn True
End Sub

Sub RHP007_Hit()
    WireRampOn True
End Sub

Sub RHP008_Hit()
    WireRampOn True
End Sub

Sub RHP009_Hit()
    WireRampOn True
End Sub

Sub RHP003_Hit
  WireRampOff
End Sub

Sub RHP004_Hit
  WireRampOff
End Sub

Sub RHP005_Hit
  WireRampOff
End Sub

Sub RHP006_Hit
  WireRampOff
End Sub


Sub RampTrigger3_Hit
    if abs(activeball.AngMomZ) > 70 then activeball.AngMomZ = 50
    activeball.AngMomZ = -abs(activeball.AngMomZ) * 3
    WireRampOff
End Sub

Sub RampTrigger4_Hit
    if abs(activeball.AngMomZ) > 70 then activeball.AngMomZ = 50
    activeball.AngMomZ = abs(activeball.AngMomZ) * 3
    WireRampOff
End Sub

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

Sub RandomSoundRampStop(obj)
  Select Case Int(rnd*3)
    Case 0: PlaySoundAtVol "wireramp_stop1", obj, 0.02*volumedial:PlaySoundAtLevelActiveBall ("Rubber_Strong_1"), Vol(ActiveBall) * RubberStrongSoundFactor * 0.6
    Case 1: PlaySoundAtVol "wireramp_stop2", obj, 0.02*volumedial:PlaySoundAtLevelActiveBall ("Rubber_Strong_2"), Vol(ActiveBall) * RubberStrongSoundFactor * 0.6
    Case 2: PlaySoundAtVol "wireramp_stop3", obj, 0.02*volumedial:PlaySoundAtLevelActiveBall ("Rubber_1_Hard"), Vol(ActiveBall) * RubberStrongSoundFactor * 0.6
  End Select
End Sub



'******************************************************
'**** END RAMP ROLLING SFX
'******************************************************

'************************************************************************
'   ZVRR: VR Room / VR Cabinet
'************************************************************************
Sub L88_animate
  If VRRoom >0 or VRtest Then
    VRCab_StartButton.blenddisableLighting = L88.GetInPlayIntensity
    VRCab_StartButtonRim.blenddisableLighting = L88.GetInPlayIntensity
  End If
End Sub

Sub LoadVRRoom
Dim VRThings
  VRCab_Blades.visible = 1
  for each VRThings in VR_Cab:VRThings.visible = 0:Next
  for each VRThings in VR_Min:VRThings.visible = 0:Next
  for each VRThings in VR_Mega:VRThings.visible = 0:Next
  for each VRThings in VR_Backglass:VRThings.visible = 0:Next

  If RenderingMode = 2 or VRTest Then
    VRRoom = VRRoomChoice
  Else
    VRRoom = 0
  End If

     If VRRoom > 0 Or Table.ShowDT Or (VRCab_BladesEnabled = 1 And Table.ShowDT = False) Then
        VRCab_Blades.visible = 1
        VRCab_Blades_OFF.visible = 1
    Else
        VRCab_Blades.visible = 0
        VRCab_Blades_OFF.visible = 0
    End If

    ' Update blade visibility based on GI state (if blades are visible)
    If VRCab_Blades.visible Then
        UpdateBladeVisibility
    End If

  If VRRoom = 1 Then
    for each VRThings in VR_Cab:VRThings.visible = 1:Next
    for each VRThings in VR_Min:VRThings.visible = 1:Next
    for each VRThings in VR_Mega:VRThings.visible = 0:Next
    for each VRThings in VR_Backglass:VRThings.visible = 1:Next
  End If

  If VRRoom = 2 Then
    for each VRThings in VR_Cab:VRThings.visible = 1:Next
    for each VRThings in VR_Min:VRThings.visible = 0:Next
    for each VRThings in VR_Mega:VRThings.visible = 0:Next
    for each VRThings in VR_Backglass:VRThings.visible = 1:Next
  End If

  If VRRoom = 3 Then
    for each VRThings in VR_Cab:VRThings.visible = 1:Next
    for each VRThings in VR_Min:VRThings.visible = 0:Next
    for each VRThings in VR_Mega:VRThings.visible = 1:Next
    for each VRThings in VR_Backglass:VRThings.visible = 1:Next
  End If

  ' Set lockbar/rails visibility
  ' VR Mode (VRRoom > 0): Always ON
  ' Desktop windowed (Table.ShowDT = True): Always ON
  ' Desktop fullscreen (Table.ShowDT = False): Controlled by option
  If VRRoom > 0 Or Table.ShowDT Then
    VRCab_LockbarRails.visible = 1 ' Always ON for VR and Desktop windowed
  Else
    VRCab_LockbarRails.visible = LockbarEnabled ' Controlled by option for Desktop fullscreen
  End If
 UpdateBackTopVisibility

End Sub



Sub ApplyBallProperties()
  Dim objBall, balls
  balls = GetBalls()

  If IsArray(balls) Then
    If UBound(balls) >= 0 Then
      For Each objBall In balls
        If IsObject(objBall) Then
          On Error Resume Next

          Select Case BallType
            Case 1: ' Golf Ball
              objBall.Image = "BlackBall"
              objBall.FrontDecal = "GolfBall"
              objBall.SphericalMapp = 1                    ' Spherical Map = OFF
              objBall.LogoMode = 1                           ' LOGO mod = OFF
              objBall.Reflection = 0.1                       ' Reflection of playfield
              objBall.BulbIntensityScale = 0.2               ' Default Bulb intensity scale

            Case 2: ' Chrome ball
              objBall.Image = "old_ass_eyes_ball"
              objBall.FrontDecal = "JPBall-Scratches"
              objBall.SphericalMapp = 1                   ' Spherical Map = ON
              objBall.LogoMode = 0                           ' LOGO mod = ON
              objBall.Reflection = 1.0                       ' Reflection of playfield
              objBall.BulbIntensityScale = 1.0               ' Default Bulb intensity scale
          End Select

          On Error Goto 0
        End If
      Next
    End If
  End If
End Sub

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
'
'        x.AddPt "Polarity", 0, 0, 0
'        x.AddPt "Polarity", 1, 0.05, - 2.7
'        x.AddPt "Polarity", 2, 0.16, - 2.7
'        x.AddPt "Polarity", 3, 0.22, - 0
'        x.AddPt "Polarity", 4, 0.25, - 0
'        x.AddPt "Polarity", 5, 0.3, - 1
'        x.AddPt "Polarity", 6, 0.4, - 2
'        x.AddPt "Polarity", 7, 0.5, - 2.7
'        x.AddPt "Polarity", 8, 0.65, - 1.8
'        x.AddPt "Polarity", 9, 0.75, - 0.5
'        x.AddPt "Polarity", 10, 0.81, - 0.5
'        x.AddPt "Polarity", 11, 0.88, 0
'        x.AddPt "Polarity", 12, 1.3, 0
'
'   x.AddPt "Velocity", 0, 0, 0.85
'   x.AddPt "Velocity", 1, 0.15, 0.85
'   x.AddPt "Velocity", 2, 0.2, 0.9
'   x.AddPt "Velocity", 3, 0.23, 0.95
'   x.AddPt "Velocity", 4, 0.41, 0.95
'   x.AddPt "Velocity", 5, 0.53, 0.95 '0.982
'   x.AddPt "Velocity", 6, 0.62, 1.0
'   x.AddPt "Velocity", 7, 0.702, 0.968
'   x.AddPt "Velocity", 8, 0.95,  0.968
'   x.AddPt "Velocity", 9, 1.03,  0.945
'   x.AddPt "Velocity", 10, 1.5,  0.945
'
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
'   x.AddPt "Polarity", 2, 0.16, - 3.7
'   x.AddPt "Polarity", 3, 0.22, - 0
'   x.AddPt "Polarity", 4, 0.25, - 0
'   x.AddPt "Polarity", 5, 0.3, - 2
'   x.AddPt "Polarity", 6, 0.4, - 3
'   x.AddPt "Polarity", 7, 0.5, - 3.7
'   x.AddPt "Polarity", 8, 0.65, - 2.3
'   x.AddPt "Polarity", 9, 0.75, - 1.5
'   x.AddPt "Polarity", 10, 0.81, - 1
'   x.AddPt "Polarity", 11, 0.88, 0
'   x.AddPt "Polarity", 12, 1.3, 0
'
'   x.AddPt "Velocity", 0, 0, 0.85
'   x.AddPt "Velocity", 1, 0.15, 0.85
'   x.AddPt "Velocity", 2, 0.2, 0.9
'   x.AddPt "Velocity", 3, 0.23, 0.95
'   x.AddPt "Velocity", 4, 0.41, 0.95
'   x.AddPt "Velocity", 5, 0.53, 0.95 '0.982
'   x.AddPt "Velocity", 6, 0.62, 1.0
'   x.AddPt "Velocity", 7, 0.702, 0.968
'   x.AddPt "Velocity", 8, 0.95,  0.968
'   x.AddPt "Velocity", 9, 1.03,  0.945
'   x.AddPt "Velocity", 10, 1.5,  0.945
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
'   x.AddPt "Polarity", 2, 0.16, - 5
'   x.AddPt "Polarity", 3, 0.22, - 0
'   x.AddPt "Polarity", 4, 0.25, - 0
'   x.AddPt "Polarity", 5, 0.3, - 2
'   x.AddPt "Polarity", 6, 0.4, - 3
'   x.AddPt "Polarity", 7, 0.5, - 4.0
'   x.AddPt "Polarity", 8, 0.7, - 3.5
'   x.AddPt "Polarity", 9, 0.75, - 3.0
'   x.AddPt "Polarity", 10, 0.8, - 2.5
'   x.AddPt "Polarity", 11, 0.85, - 2.0
'   x.AddPt "Polarity", 12, 0.9, - 1.5
'   x.AddPt "Polarity", 13, 0.95, - 1.0
'   x.AddPt "Polarity", 14, 1, - 0.5
'   x.AddPt "Polarity", 15, 1.1, 0
'   x.AddPt "Polarity", 16, 1.3, 0
'
'   x.AddPt "Velocity", 0, 0, 0.85
'   x.AddPt "Velocity", 1, 0.15, 0.85
'   x.AddPt "Velocity", 2, 0.2, 0.9
'   x.AddPt "Velocity", 3, 0.23, 0.95
'   x.AddPt "Velocity", 4, 0.41, 0.95
'   x.AddPt "Velocity", 5, 0.53, 0.95 '0.982
'   x.AddPt "Velocity", 6, 0.62, 1.0
'   x.AddPt "Velocity", 7, 0.702, 0.968
'   x.AddPt "Velocity", 8, 0.95,  0.968
'   x.AddPt "Velocity", 9, 1.03,  0.945
'   x.AddPt "Velocity", 10, 1.5,  0.945

' Next
'
' ' SetObjects arguments: 1: name of object 2: flipper object: 3: Trigger object around flipper
' LF.SetObjects "LF", LeftFlipper, TriggerLF
' RF.SetObjects "RF", RightFlipper, TriggerRF
'End Sub

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

  Public Sub ReProcessBalls(aBall) 'save data of balls in flipper range
    If FlipperOn() Then
      Dim x
      For x = 0 To UBound(balls)
        If Not IsEmpty(balls(x)) Then
          if balls(x).ID = aBall.ID Then
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

  Public Sub PolarityCorrect(aBall)
    If FlipperOn() Then
      Dim tmp, BallPos, x, IDX, Ycoef, BalltoFlip, BalltoBase, NoCorrection, checkHit
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
          BalltoFlip = DistanceFromFlipperAngle(BallData(x).x, BallData(x).y, Flipper, FlipStartAngle)
          If ballpos > 0.65 Then  Ycoef = LinearEnvelope(BallData(x).Y, YcoefIn, YcoefOut)                'find safety coefficient 'ycoef' data
        End If
      Next

      If BallPos = 0 Then 'no ball data meaning the ball is entering and exiting pretty close to the same position, use current values.
        BallPos = PSlope(aBall.x, FlipperStart, 0, FlipperEnd, 1)
        If ballpos > 0.65 Then  Ycoef = LinearEnvelope(aBall.Y, YcoefIn, YcoefOut)                        'find safety coefficient 'ycoef' data
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

        If Enabled Then aBall.Velx = aBall.Velx*VelCoef
        If Enabled Then aBall.Vely = aBall.Vely*VelCoef
      End If

      'Polarity Correction (optional now)
      If Not IsEmpty(PolarityIn(0) ) Then
        Dim AddX
        AddX = LinearEnvelope(BallPos, PolarityIn, PolarityOut) * LR

        If Enabled and NoCorrection = 0 Then aBall.VelX = aBall.VelX + 1 * (AddX*ycoef*PartialFlipcoef*VelCoef)
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
'  Const EOSReturn = 0.045  'late 70's to mid 80's
' Const EOSReturn = 0.035  'mid 80's to early 90's
   Const EOSReturn = 0.025  'mid 90's and later

LFEndAngle = Leftflipper.endangle
RFEndAngle = RightFlipper.endangle

Sub FlipperActivate(Flipper, FlipperPress)
  FlipperPress = 1
  Flipper.Elasticity = FElasticity

  Flipper.eostorque = EOST
  Flipper.eostorqueangle = EOSA
End Sub

'Sub FlipperDeactivate(Flipper, FlipperPress)
' FlipperPress = 0
' Flipper.eostorqueangle = EOSA
' Flipper.eostorque = EOST * EOSReturn / FReturn

' If Abs(Flipper.currentangle) <= Abs(Flipper.endangle) + 0.1 Then
'   Dim b', BOT
    '   BOT = GetBalls

'   For b = 0 To UBound(gBOT)
'     If Distance(gBOT(b).x, gBOT(b).y, Flipper.x, Flipper.y) < 55 Then 'check for cradle
'       If gBOT(b).vely >= - 0.4 Then gBOT(b).vely =  - 0.4
'     End If
'   Next
' End If
'End Sub

'Sub FlipperDeactivate(Flipper, FlipperPress)
' FlipperPress = 0
' Flipper.eostorqueangle = EOSA
' Flipper.eostorque = EOST * EOSReturn / FReturn
'
' If Abs(Flipper.currentangle) <= Abs(Flipper.endangle) + 0.1 Then
'   Dim b', BOT
'   '   BOT = GetBalls   ' if you use LF.ReProcessBalls / gBOT elsewhere, this can stay commented
'
'   For b = 0 To UBound(gBOT)
'           ' make sure this slot actually holds a ball object
'            If Not (gBOT(b) Is Nothing) Then
'         If Distance(gBOT(b).x, gBOT(b).y, Flipper.x, Flipper.y) < 55 Then 'check for cradle
'           If gBOT(b).vely >= -0.4 Then gBOT(b).vely = -0.4
'         End If
'            End If
'   Next
' End If
'End Sub


Sub FlipperDeactivate(Flipper, FlipperPress)
    Dim BOT, b

    FlipperPress = 0
    Flipper.eostorqueangle = EOSA
    Flipper.eostorque = EOST * EOSReturn / FReturn

    If Abs(Flipper.currentangle) <= Abs(Flipper.endangle) + 0.1 Then
        BOT = GetBalls   ' fresh list of valid balls on the table

        If IsArray(BOT) Then
            For b = 0 To UBound(BOT)
                If IsObject(BOT(b)) Then
                    ' check for cradle near this flipper
                    If Distance(BOT(b).x, BOT(b).y, Flipper.x, Flipper.y) < 55 Then
                        ' clamp downward speed a bit
                        If BOT(b).VelY >= -0.4 Then BOT(b).VelY = -0.4
                    End If
                End If
            Next
        End If
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
' ZSSC: SLINGSHOT CORRECTION FUNCTIONS by apophis
'******************************************************
' To add these slingshot corrections:
'  - On the table, add the endpoint primitives that define the two ends of the Slingshot
'  - Initialize the SlingshotCorrection objects in InitSlingCorrection
'  - Call the .VelocityCorrect methods from the respective _Slingshot event sub


Dim LS: Set LS = New SlingshotCorrection
Dim RS: Set RS = New SlingshotCorrection
'Dim ULS: Set ULS = New SlingshotCorrection
'Dim URS: Set URS = New SlingshotCorrection

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

' The following sub are needed, however they exist in the ZMAT maths section of the script. Uncomment below if needed
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

'Flipper collide subs
Sub LeftFlipper_Collide(parm)
  CheckLiveCatch Activeball, LeftFlipper, LFCount, parm
  LF.ReProcessBalls ActiveBall
  LeftFlipperCollide parm
End Sub

Sub RightFlipper_Collide(parm)
  CheckLiveCatch Activeball, RightFlipper, RFCount, parm
  RF.ReProcessBalls ActiveBall
  RightFlipperCollide parm
End Sub

Sub UpperFlipper_Collide(parm)
  RightFlipperCollide parm
End Sub


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
    aBall.velz = aBall.velz * coef
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
'   ZBOU: VPW TargetBouncer for targets and posts by Iaakki, Wrd1972, Apophis
'******************************************************

Const TargetBouncerEnabled = 1    '0 = normal standup targets, 1 = bouncy targets
Const TargetBouncerFactor = 0.9   'Level of bounces. Recommmended value of 0.7-1.0

sub TargetBouncer(aBall,defvalue)
    dim zMultiplier, vel, vratio
    if TargetBouncerEnabled = 1 and aball.z < 30 then
        'debug.print "velx: " & aball.velx & " vely: " & aball.vely & " velz: " & aball.velz
        vel = BallSpeed(aBall)
        if aBall.velx = 0 then vratio = 1 else vratio = aBall.vely/aBall.velx
        Select Case Int(Rnd * 6) + 1
            Case 1: zMultiplier = 0.2*defvalue
      Case 2: zMultiplier = 0.25*defvalue
            Case 3: zMultiplier = 0.3*defvalue
      Case 4: zMultiplier = 0.4*defvalue
            Case 5: zMultiplier = 0.45*defvalue
            Case 6: zMultiplier = 0.5*defvalue
        End Select
        aBall.velz = abs(vel * zMultiplier * TargetBouncerFactor)
        aBall.velx = sgn(aBall.velx) * sqr(abs((vel^2 - aBall.velz^2)/(1+vratio^2)))
        aBall.vely = aBall.velx * vratio
        'debug.print "---> velx: " & aball.velx & " vely: " & aball.vely & " velz: " & aball.velz
        'debug.print "conservation check: " & BallSpeed(aBall)/vel
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

Sub soundtrigger_Hit()
    If ActiveBall.VelY < 0 Then      ' ball moving UP the table
    PlaySoundAtLevelStatic ("WarehouseKick"), DrainSoundLevel, Kickback
    End If
End Sub

Sub collidesound_Hit()
    If ActiveBall.VelY < 0 Then      ' ball moving UP the table
'PlaySound "Saucer_Enter_2"
  Select Case Int(Rnd * 7) + 1
    Case 1
      PlaySoundAtLevelStatic SoundFX("Ball_Collide_1", DOFContactors), SaucerKickSoundLevel, CollideSound
    Case 2
      PlaySoundAtLevelStatic SoundFX("Ball_Collide_2", DOFContactors), SaucerKickSoundLevel, CollideSound
    Case 3
      PlaySoundAtLevelStatic SoundFX("Ball_Collide_3", DOFContactors), SaucerKickSoundLevel, CollideSound
    Case 4
      PlaySoundAtLevelStatic SoundFX("Ball_Collide_4", DOFContactors), SaucerKickSoundLevel, CollideSound
    Case 5
      PlaySoundAtLevelStatic SoundFX("Ball_Collide_5", DOFContactors), SaucerKickSoundLevel, CollideSound
    Case 6
      PlaySoundAtLevelStatic SoundFX("Ball_Collide_6", DOFContactors), SaucerKickSoundLevel, CollideSound
    Case 7
      PlaySoundAtLevelStatic SoundFX("Ball_Collide_7", DOFContactors), SaucerKickSoundLevel, CollideSound
  End Select
  End if
End Sub

Sub RampTrigger1_hit
if lramp.collidable=1 then
  If activeball.vely < 0 Then
    PlaySoundAtLevelStatic SoundFX("rampshort1", DOFContactors), 0.2, Ramptrigger1
  Else
    Stopsound "rampshort1"
  End If
end if
end Sub

Sub RampTrigger1b_hit
  If activeball.vely < 0 Then
    Stopsound "rampshort1"
  Else
    PlaySoundAtLevelStatic SoundFX("rampshort1", DOFContactors), 0.2, Ramptrigger1
  End If
end Sub

Sub RampTrigger1c_hit
    Stopsound "rampshort1"
end Sub

Sub RampTrigger2_hit
if rramp.collidable=1 then
  If activeball.vely < 0 Then
    PlaySoundAtLevelStatic SoundFX("rampshort2", DOFContactors), 0.2, RampTrigger2
  Else
    Stopsound "rampshort2"
  End If
end if
end Sub

Sub RampTrigger2b_hit
  If activeball.vely < 0 Then
    Stopsound "rampshort2"
  Else
    PlaySoundAtLevelStatic SoundFX("rampshort2", DOFContactors), 0.2, Ramptrigger2
  End If
end Sub

Sub RampTrigger2c_hit
    Stopsound "rampshort2"
end Sub


Sub Plunger_Init()

End Sub

' Thalamus : Exit in a clean and proper way
Sub Table_exit
  Controller.Pause = False
  Controller.Stop
End Sub



  Option Explicit
   Randomize

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

'******************* Options *********************
' DMD/BAckglass Setting
Const LampShader        = 1         'Set to 1 to use high performance lamp shaders, set to 0 if you see poor performance
Const HighPerformance     = 1         'Set to 1 for high resolution texture swapping, set to 0 if you see stutter and don't need this feature
Const BallSize = 50
Const BallMass = 1.0

Dim VarHidden, UseVPMDMD

If Table.ShowDT = true then
  UseVPMDMD = True
  VarHidden = 1
  l40a.Visible = false
  l40b.Visible = false
  l40a1.Visible = true
  l40b1.Visible = true
  Ramp19.Visible = true
  Ramp18.Visible = true
  SideWood.Visible = True
  lockdown.Visible = true
else
  UseVPMDMD = False
  VarHidden = 0
  l40a1.Visible = false
  l40b1.Visible = false
  l40a.Visible = true
  l40b.Visible = true
  Ramp19.Visible = false
  Ramp18.Visible = false
  'SideWood.Visible = false
  lockdown.Visible = false
end if

LoadVPM "01130000","CAPCOM.VBS",3.10

  Const FlipSolNudge = 1      ' Make flipper solenoid to alter active ball behavior
  Const LiveCatch = 24      ' Live catch window size. Higher value will make live catching easier. 8-32
  Const slackyflips = 1     ' Visible Flipper slack on hard ball collision





'********************
'Standard definitions
'********************

  Const cGameName = "kpv106" 'change the romname here

     Const UseSolenoids = 2
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
    Dim x
    Dim Bump1,Bump2,Bump3,Mech3bank,bsTrough,bsVUK,visibleLock,bsTEject,bsSVUK,bsRScoop,bsLock
  Dim dtDropL, dtDropR
  Dim PlungerIM
  Dim PMag
  Dim mag2
  Dim bsRHole
  Dim FireButtonFlag:FireButtonFlag = 0
  Dim IsStarted:IsStarted = 0

  Sub Table_Init

'*****

Kicker1.CreateBall
Kicker1.Kick 0,0
Kicker1.Enabled = false

Kicker2.CreateBall
Kicker2.Kick 0,0
Kicker2.Enabled = false

Controller.Switch(47) = 1 'ramp down active
' Thalamus : Was missing 'vpminit me'
  vpminit me


  With Controller
    .GameName = cGameName
    If Err Then MsgBox "Can't start Game " & cGameName & vbNewLine & Err.Description:Exit Sub
    .SplashInfoLine = "Kingpin"
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

    Const IMPowerSetting = 65
    Const IMTime = 0.6
    Set plungerIM = New cvpmImpulseP
    With plungerIM
        .InitImpulseP swplunger, IMPowerSetting, IMTime
        .Random 1
    .Switch 43
        .InitExitSnd SoundFX("plunger2",DOFContactors), SoundFX("plunger",DOFContactors)
        .CreateEvents "plungerIM"
    End With

'**Trough
    Set bsTrough = New cvpmBallStack
    bsTrough.InitSw 35,36,37,38,39,0,0,0
    bsTrough.InitKick BallRelease, 50, 10
    bsTrough.InitExitSnd SoundFX("ballrel",DOFContactors), SoundFX("Solenoid",DOFContactors)

    bsTrough.Balls = 4

  set dtDropL=new cvpmDropTarget
  dtDropL.InitDrop Array(sw25,sw26,sw27,sw28),Array(25, 26, 27, 28)
  dtDropL.InitSnd SoundFX("DTR",DOFContactors),SoundFX("DTResetL",DOFContactors)

  set dtDropR=new cvpmDropTarget
  dtDropR.InitDrop Array(sw29,sw30,sw31),Array(29, 30, 31)
  dtDropR.InitSnd SoundFX("DTR",DOFContactors),SoundFX("DTResetR",DOFContactors)

  Set bsVUK=New cvpmBallStack
  bsVUK.InitSw 0,51,0,0,0,0,0,0
  bsVUK.InitKick sw51b,173,28
  bsVUK.KickForceVar = 3
  bsVUK.KickAngleVar = 3
  bsVUK.KickZ = 55
  bsVUK.KickBalls = 1
  bsVUK.InitExitSnd SoundFX("scoopexit",DOFContactors), SoundFX("rail",DOFContactors)

  Set bsLock=New cvpmBallStack
  bsLock.InitSw 0,44,45,46,0,0,0,0
  bsLock.InitKick GunKicker,170,10
  bsLock.KickForceVar = 3
  bsLock.KickAngleVar = 3
  bsLock.InitExitSnd SoundFX("scoopexit",DOFContactors), SoundFX("rail",DOFContactors)

'**Nudging
      vpmNudge.TiltSwitch=-7
    vpmNudge.Sensitivity=1
    vpmNudge.TiltObj=Array(Bumper1,Bumper2,Bumper3,LeftSlingshot,RightSlingshot)

      '**Main Timer init
  PinMAMETimer.Interval = PinMAMEInterval
  PinMAMETimer.Enabled = 1

  End Sub

'*****Keys
Sub Table_KeyDown(ByVal Keycode)
  'nFozzy Begin'
  If keycode = LeftFlipperKey Then LFPress = 1 : DOF 101,1
  If keycode = RightFlipperKey Then rfpress = 1 : DOF 102,1
  'nFozzy End'

  If Keycode = LeftFlipperKey then
    Controller.Switch(5)=1
    If bsTrough.balls < 4 and FlipperDisabled < 4 Then
'     PlaySound SoundFX("flipperupleft",DOFContactors)
      PlaySoundAtVol SoundFX("Flipper_L" & int(rnd(1)*11)+1 , DOFContactors), LeftFlipper, 1
      LF.Fire ' LeftFlipper.RotateToEnd
      FlipperDisabled = FlipperDisabled + 1
    End If
    Exit Sub
  End If
  If Keycode = RightFlipperKey then
    Controller.Switch(6)=1
    If bsTrough.balls < 4 and FlipperDisabled < 4 Then
'     PlaySound SoundFX("flipperupright",DOFContactors)
      PlaySoundAtVol SoundFX("Flipper_R" & int(rnd(1)*11)+1 , DOFContactors), RightFlipper, 1

      RF.fire ' RightFlipper.RotateToEnd
      FlipperDisabled = FlipperDisabled + 1
    End If
    Exit Sub
  End If

    If keycode = PlungerKey Then vpmTimer.PulseSw 14
    If vpmKeyDown(keycode) Then Exit Sub
End Sub


Sub Table_KeyUp(ByVal keycode)
  If vpmKeyUp(keycode) Then Exit Sub
    'nFozzy Begin'
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
  'nFozzy End'
  If Keycode = LeftFlipperKey then
    Controller.Switch(5)=0
'   PlaySound SoundFX("flipperdown",DOFContactors)
    PlaySoundAtVol SoundFX("Flipper_Left_Down_" & int(rnd(1)*7)+1 , DOFContactors), LeftFlipper, 1

    FlipperDeActivate LeftFlipper, LFPress
    LeftFlipper.RotateToStart
    Exit Sub
  End If
  If Keycode = RightFlipperKey then
    Controller.Switch(6)=0
'   PlaySound SoundFX("flipperdown",DOFContactors)
    PlaySoundAtVol SoundFX("Flipper_Right_Down_" & int(rnd(1)*8)+1 , DOFContactors), RightFlipper, 1
    FlipperDeActivate RightFlipper, RFPress
    RightFlipper.RotateToStart
    Exit Sub
  End If

  If Keycode = StartGameKey Then Controller.Switch(7) = 0 and IsStarted = 1
End Sub

   'Solenoids
SolCallback(1)  = "bsTrough.SolIn"
SolCallback(2)  = "bsTrough.SolOut"

SolCallback(6)    = "dtDropL.SolDropUp"
SolCallback(7)    = "dtDropR.SolDropUp"
SolCallback(8) = "bsLock.SolOut" 'gun eject
SolCallback(9) = "SolLFlipper"
SolCallback(10) = "SolRFlipper"
SolCallback(11) = "bsVUK.SolOut" 'slot machine eject
SolCallback(12) = "SlotMachineMotor"
SolCallBack(13) = "LoopGate"
SolCallback(14) = "SolLeftRamp"

SolCallback(18) = "LeftRampFlash"
SolCallback(19) = "f19.State="
SolCallback(21) = "f21.State="
SolCallback(22) = "f22.State="
SolCallback(23) = "RightRampFlash"
SolCallback(24) = "f24.State="
SolCallback(25) = "f25.State="
SolCallback(26) = "f26.State="
SolCallback(27) = "SetLamp 131,"
SolCallback(28) = "f28.State="
SolCallback(29) = "f29.State="
SolCallback(30) = "f30.State="
SolCallback(31) = "f31.State="
SolCallback(32) = "SolAutoFire"

' SolCallback(sLRFlipper) = "SolRFlipper"
' SolCallback(sLLFlipper) = "SolLFlipper"

Sub start2_timer
  IsStarted = 1
  Controller.Switch(9) = 0
  start2.enabled = 0
End Sub

Sub SlotMachineMotor(Enabled)
  If Enabled then
    SlotTimer.Enabled = 1
    If IsStarted = 0 then Controller.Switch(9) = 1:start2.enabled = 1:End If '***ENABLE THIS IS LAMPS FLICKER (AND YOU REFUSE TO UPDATE VPINMAME)
  Else
    SlotTimer.Enabled = 0
  End If
End Sub



'************* SLOT MACHINE *************
'SLOT MACHINE BASED ON SCRIPT BY DESTRUK AND UNCLE REAMUS
'Row 1=Money 320
'Row 2=Goods 0
'Row 3=Sevens 40
'Row 4=Gangsters 80
'Row 5=Bars 120
'Row 6=Power 160
'Row 7=Guns 200
'Row 8=Crazy Cash 240
'Row 9=Cherries 280

Dim SlotPos
SlotPos=0

Sub SlotTimer_Timer
SlotPos=SlotPos+1

Select Case SlotPos
  Case 0:Controller.Switch(52)=1 'goods
  Case 39:Controller.Switch(52)=0
  Case 40:Controller.Switch(52)=1 'sevens
  Case 79:Controller.Switch(52)=0
  Case 80:Controller.Switch(52)=1 'gangsters
  Case 119:Controller.Switch(52)=0
  Case 120:Controller.Switch(52)=1 'bars
  Case 159:Controller.Switch(52)=0
  Case 160:Controller.Switch(52)=1 'power
  Case 199:Controller.Switch(52)=0
  Case 200:Controller.Switch(52)=1 'guns
  Case 239:Controller.Switch(52)=0
  Case 240:Controller.Switch(52)=1 'crazy cash
  Case 279:Controller.Switch(52)=0
  Case 280:Controller.Switch(52)=1 'cherries
  Case 319:Controller.Switch(52)=0
  Case 320:Controller.Switch(52)=1 'money
  Case 359:Controller.Switch(52)=0
  Case 360:SlotPos = 0
End Select
End Sub

'''Flashers'''

Dim RRFON:RRFON = 0
Dim LRFON:LRFON = 0
Sub LeftRampFlash(Enabled)
  If Enabled Then
    SetFlash 130, 1
    SetLamp 140, 1
    SetLamp 132, 1
    SetLamp 133, 1
    SetLamp 134, 1
    SetLamp 136, 1
    SetLamp 130, 1
    SetLamp 141, 1
    SetLamp 142, 1
    SetLamp 143, 1
    p122r.image = "Wire2R_ON3"
    If LRFON = 1 then
      SetLamp139, 1
    End If
    If LRFON = 0 then
      SetLamp 137, 1
    End If
  Else
    SetFlash 130, 0
    SetLamp 140, 0
    SetLamp 132, 0
    SetLamp 133, 0
    SetLamp 134, 0
    SetLamp 136, 0
    SetLamp 130, 0
    SetLamp 141, 0
    SetLamp 142, 0
    SetLamp 143, 0
    p122r.image = "Off"
    If LRFON = 1 then
      SetLamp138, 1
    End If
    If LRFON = 0 then
      SetLamp 137, 0
    End If
  End If
End Sub

Sub RightRampFlash(Enabled)
  If Enabled Then
    SetFlash 129, 1
    SetLamp 129, 1
    SetLamp 135, 1
    SetLamp 144, 1
    If RRFON = 1 then
      SetLamp139, 1
    End If
    If RRFON = 0 then
      SetLamp 138, 1
    End If
  Else
    SetFlash 129, 0
    SetLamp 129, 0
    SetLamp 135, 0
    SetLamp 144, 0
    If RRFON = 1 then
      SetLamp137, 1
    End If
    If RRFON = 0 then
      SetLamp 138, 0
    End If
  End If
End Sub

Dim LoopOpen:LoopOpen = 0
Sub LoopGate(Enabled)
  If Enabled then
    If LoopOpen = 1 then
      Gate2.Collidable = 1
      LoopOpen = 0
    Else
      Gate2.Collidable = 0
      LoopOpen = 1
    End If
  End If
End Sub

dim RampUp:RampUp = 0
Sub SolLeftRamp(Enabled)
  If Enabled Then
    If RampUp = 0 Then
      LeftRampFlipper.RotateToEnd
      Controller.Switch(47) = 0
      RampUp = 1
    Else
      LeftRampFlipper.RotateToStart
      Controller.Switch(47) = 1
      RampUp = 0
    End If
  End If

End Sub



Dim KickNow:KickNow = 0
'Sub Kicker3_Hit:Kicker3.DestroyBall:GunEject.Enabled = 1:GunEjectFlipper.RotateToEnd:End Sub
Sub Kicker4_Hit:GunEject.Enabled = 1:GunEjectFlipper.RotateToEnd:End Sub

Sub GunEject_Timer
  'Kicker4.CreateBall
  Kicker4.KickZ 170, 10, 0, 55
  GunEjectFlipper.RotateToStart
  GunEject.Enabled = 0
End Sub



Sub solTrough(Enabled)
  If Enabled Then
    bsTrough.ExitSol_On
  End If
 End Sub

Sub solAutofire(Enabled)
  If Enabled Then
    PlungerIM.AutoFire
  End If
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
  Primitive_RXRspinner1.RotX = -(sw61.currentangle) +90
    Primitive_RXRspinner2.RotX = -(sw17.currentangle) +90
  If RampUp = 1 then Ramp5.Collidable = 0
  If RampUp = 0 then Ramp5.Collidable = 1
SLOTmachineCylinder.RotX = SlotPos
  Primitive_Metalramp1.ObjRotX = LeftRampFlipper.CurrentAngle
  gunejectprim.ObjRotX = GunEjectFlipper.CurrentAngle
End Sub



'******************************************
' Use FlipperTimers to call div subs
'******************************************




Dim FlipperDisabled
       Sub SolLFlipper(Enabled)
           If Enabled Then
        FlipperDisabled = 0
       Else
        If bsTrough.balls = 4 Then
          Debug.print "soloff"
            'PlaySound "flipperdown"
          LeftFlipper.RotateToStart
        end if
           End If
       End Sub

       Sub SolRFlipper(Enabled)
           If Enabled Then
        FlipperDisabled = 0
       Else
        If bsTrough.balls = 4 Then
          Debug.print "solof"
            'PlaySound "flipperdown"
          RightFlipper.RotateToStart
        end if

           End If
       End Sub



'***Slings and rubbers

 Sub LeftSlingShot_Slingshot
  Leftsling = True
  Controller.Switch(41) = 1
  PlaySoundAtVol Soundfx("Sling_L" & int(rnd(1)*10)+1 , DOFContactors), Left1, 1 :LeftSlingshot.TimerEnabled = 1:f50.visible = 1

  End Sub

Dim Leftsling:Leftsling = False

Sub LS_Timer()
  If Leftsling = True and Left1.ObjRotZ < -7 then Left1.ObjRotZ = Left1.ObjRotZ + 2
  If Leftsling = False and Left1.ObjRotZ > -20 then Left1.ObjRotZ = Left1.ObjRotZ - 2
  If Left1.ObjRotZ >= -7 then Leftsling = False
  If Leftsling = True and Left2.ObjRotZ > -212.5 then Left2.ObjRotZ = Left2.ObjRotZ - 2
  If Leftsling = False and Left2.ObjRotZ < -199.5 then Left2.ObjRotZ = Left2.ObjRotZ + 2
  If Left2.ObjRotZ <= -212.5 then Leftsling = False
  If Leftsling = True and Left3.TransZ > -23 then Left3.TransZ = Left3.TransZ - 4
  If Leftsling = False and Left3.TransZ < -0 then Left3.TransZ = Left3.TransZ + 4
  If Left3.TransZ <= -23 then Leftsling = False
End Sub

 Sub LeftSlingShot_Timer:Me.TimerEnabled = 0:Controller.Switch(41) = 0:f50.visible = 0:End Sub

 Sub RightSlingShot_Slingshot
  Rightsling = True
  Controller.Switch(42) = 1
  PlaySoundAtVol Soundfx("Sling_L" & int(rnd(1)*8)+1 , DOFContactors), Right1, 1:RightSlingshot.TimerEnabled = 1:f51.visible = 1
  End Sub

 Dim Rightsling:Rightsling = False

Sub RS_Timer()
  If Rightsling = True and Right1.ObjRotZ > 7 then Right1.ObjRotZ = Right1.ObjRotZ - 2
  If Rightsling = False and Right1.ObjRotZ < 20 then Right1.ObjRotZ = Right1.ObjRotZ + 2
  If Right1.ObjRotZ <= 7 then Rightsling = False
  If Rightsling = True and Right2.ObjRotZ < 212.5 then Right2.ObjRotZ = Right2.ObjRotZ + 2
  If Rightsling = False and Right2.ObjRotZ > 199.5 then Right2.ObjRotZ = Right2.ObjRotZ - 2
  If Right2.ObjRotZ >= 212.5 then Rightsling = False
  If Rightsling = True and Right3.TransZ > -23 then Right3.TransZ = Right3.TransZ - 4
  If Rightsling = False and Right3.TransZ < -0 then Right3.TransZ = Right3.TransZ + 4
  If Right3.TransZ <= -23 then Rightsling = False
End Sub

 Sub RightSlingShot_Timer:Me.TimerEnabled = 0:Controller.Switch(42) = 0:f51.visible = 0:End Sub

Sub Bumper1_Hit:vpmTimer.PulseSw 57: PlaySoundAtVol SoundFX("Bumpers_Top_" & int(rnd(1)*5)+1 , DOFContactors), ActiveBall, 1:Bumper1.TimerEnabled = 1:End Sub
Sub Bumper1_Timer:Bumper1.TimerEnabled = 0:End Sub

Sub Bumper2_Hit:vpmTimer.PulseSw 58:PlaySoundAtVol SoundFX("Bumpers_Top_" & int(rnd(1)*5)+1 , DOFContactors), ActiveBall, 1:Bumper2.TimerEnabled = 1:End Sub
Sub Bumper2_Timer:Bumper2.TimerEnabled = 0:End Sub

Sub Bumper3_Hit:vpmTimer.PulseSw 59:PlaySoundAtVol SoundFX("Bumpers_Top_" & int(rnd(1)*5)+1 , DOFContactors), ActiveBall, 1:Bumper3.TimerEnabled = 1:End Sub
Sub Bumper3_Timer:Bumper3.TimerEnabled = 0:End Sub

 'Drains and Kickers
Dim BallInPlay:BallInPlay = 0

Sub Drain_Hit
  PlaySound ("Drain_" & int(rnd(1)*11)+1 ), 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioPan(ActiveBall)
' PlaySound "Drain"
  bsTrough.AddBall Me
  BallInPlay = BallInPlay - 1
End Sub



Sub BallRelease_UnHit(): BallInPlay = BallInPlay + 1: PlaySoundAtVol SoundFX("Ballrelease" & int(rnd(1)*7)+1 , DOFContactors), BallRelease, 1 : End Sub





'Sub sw44_Hit:bsLock.AddBall Me:End Sub

Sub sw44_Hit
  Set aBall = ActiveBall
  aZpos = 50
  Me.TimerInterval = 2
  Me.TimerEnabled = 1
End Sub



Sub sw44_Timer
  aBall.Z = aZpos
  aZpos = aZpos-2
  If aZpos <40 Then
    Me.TimerEnabled = 0
    Me.DestroyBall
    bsLock.AddBall Me
  End If
End Sub

Sub sw51a_Hit:sw51.Enabled = 1:End Sub
'Sub sw51_Hit:bsVUK.AddBall Me:sw51.Enabled = 0:End Sub
Dim aBall, aZpos



Sub sw51_Hit
  Set aBall = ActiveBall
  aZpos = 50
  Me.TimerInterval = 2
  Me.TimerEnabled = 1
  playsound "scoopleft"
End Sub



Sub sw51_Timer
  aBall.Z = aZpos
  aZpos = aZpos-2
  If aZpos <40 Then
    Me.TimerEnabled = 0
    Me.DestroyBall
    bsVUK.AddBall Me
    sw51.Enabled = 0
  End If
End Sub

Sub sw19_Hit:Me.TimerEnabled = 1:Controller.Switch(19) = 1:PlaySoundAtVol ("Rollover_" & int(rnd(1)*4)+1), ActiveBall, 1:End Sub 'left inlane
Sub sw19_Timer:Me.TimerEnabled = 0:Controller.Switch(19) = 0:End Sub
Sub sw20_Hit:Me.TimerEnabled = 1:Controller.Switch(20) = 1:PlaySoundAtVol ("Rollover_" & int(rnd(1)*4)+1), ActiveBall, 1:End Sub 'right inlane
Sub sw20_Timer:Me.TimerEnabled = 0:Controller.Switch(20) = 0:End Sub
Sub sw21_Hit:Me.TimerEnabled = 1:Controller.Switch(21) = 1:PlaySoundAtVol ("Rollover_" & int(rnd(1)*4)+1), ActiveBall, 1:End Sub 'left outlane
Sub sw21_Timer:Me.TimerEnabled = 0:Controller.Switch(21) = 0:End Sub
Sub sw22_Hit:Me.TimerEnabled = 1:Controller.Switch(22) = 1:PlaySoundAtVol ("Rollover_" & int(rnd(1)*4)+1), ActiveBall, 1:End Sub 'right outlane
Sub sw22_Timer:Me.TimerEnabled = 0:Controller.Switch(22) = 0:End Sub
Sub sw23_Hit:Me.TimerEnabled = 1:Controller.Switch(23) = 1:PlaySoundAtVol ("Rollover_" & int(rnd(1)*4)+1), ActiveBall, 1:End Sub 'left orbit
Sub sw23_Timer:Me.TimerEnabled = 0:Controller.Switch(23) = 0:End Sub
Sub sw24_Hit:Me.TimerEnabled = 1:Controller.Switch(24) = 1:PlaySoundAtVol ("Rollover_" & int(rnd(1)*4)+1), ActiveBall, 1:End Sub 'right orbit
Sub sw24_Timer:Me.TimerEnabled = 0:Controller.Switch(24) = 0:End Sub

Sub sw25_Hit:dtDropL.Hit 1:End Sub  'K
Sub sw26_Hit:dtDropL.Hit 2:End Sub  'I
Sub sw27_Hit:dtDropL.Hit 3:End Sub  'N
Sub sw28_Hit:dtDropL.Hit 4:End Sub  'G
Sub sw29_Hit:dtDropR.Hit 1:End Sub  'P
Sub sw30_Hit:dtDropR.Hit 2:End Sub  'I
Sub sw31_Hit:dtDropR.Hit 3:End Sub  'N

Sub sw32_Hit  : vpmTimer.PulseSw 49:sw32.TimerEnabled = 1:sw32p.TransX = -4: playsoundAtVol SoundFX("target",DOFContactors), ActiveBall, 1: End Sub 'captive ball
Sub sw32_Timer:Me.TimerEnabled = 0:sw32p.TransX = 0:End Sub

Sub swPlunger_Hit:Me.TimerEnabled = 1:Controller.Switch(43) = 1:PlaySoundAtVol ("Rollover_" & int(rnd(1)*4)+1), ActiveBall, 1:End Sub 'plunger lane
Sub swPlunger_Timer:Me.TimerEnabled = 0:Controller.Switch(43) = 0:End Sub

Sub sw53_Hit:Me.TimerEnabled = 1:Controller.Switch(53) = 1: PlaySoundAtVol ("Rollover_" & int(rnd(1)*4)+1), ActiveBall, 1:End Sub 'k top rollover
Sub sw53_Timer:Me.TimerEnabled = 0:Controller.Switch(53) = 0:End Sub
Sub sw54_Hit:Me.TimerEnabled = 1:Controller.Switch(54) = 1:PlaySoundAtVol ("Rollover_" & int(rnd(1)*4)+1), ActiveBall, 1:End Sub 'i top rollover
Sub sw54_Timer:Me.TimerEnabled = 0:Controller.Switch(54) = 0:End Sub
Sub sw55_Hit:Me.TimerEnabled = 1:Controller.Switch(55) = 1:PlaySoundAtVol ("Rollover_" & int(rnd(1)*4)+1), ActiveBall, 1:End Sub 'd top rollover
Sub sw55_Timer:Me.TimerEnabled = 0:Controller.Switch(55) = 0:End Sub

Sub sw49_Hit  : vpmTimer.PulseSw 49:sw49.TimerEnabled = 1:sw49p.TransX = -4: playsoundAtVol SoundFX("target",DOFContactors), ActiveBall, 1: End Sub
Sub sw49_Timer:Me.TimerEnabled = 0:sw49p.TransX = 0:End Sub
Sub sw50_Hit  : vpmTimer.PulseSw 50:sw50.TimerEnabled = 1:sw50p.TransX = -4: playsoundAtVol SoundFX("target",DOFContactors), ActiveBall, 1: End Sub
Sub sw50_Timer:Me.TimerEnabled = 0:sw50p.TransX = 0:End Sub
Sub sw60_Hit  : vpmTimer.PulseSw 60:sw60.TimerEnabled = 1:sw60p.TransX = -4: playsoundAtVol SoundFX("target",DOFContactors), ActiveBall, 1: End Sub
Sub sw60_Timer:Me.TimerEnabled = 0:sw60p.TransX = 0:End Sub
Sub sw63_Hit  : vpmTimer.PulseSw 63:sw63.TimerEnabled = 1:sw63p.TransX = -4: playsoundAtVol SoundFX("target",DOFContactors), ActiveBall, 1: End Sub
Sub sw63_Timer:Me.TimerEnabled = 0:sw63p.TransX = 0:End Sub

' Thalamus : This sub is used twice - this means ... this one WAS NOT USED
' Sub sw61_Spin:vpmTimer.PulseSw 61:PlaySound "spinner":End Sub 'left ramp spinner

Sub sw17_Spin:vpmTimer.PulseSw 17:PlaySoundAtVol "spinner", sw17, 1:End Sub 'right ramp spinner
Sub sw18_Hit:Me.TimerEnabled = 1:Controller.Switch(18) = 1:End Sub
Sub sw18_Timer:Me.TimerEnabled = 0:Controller.Switch(18) = 0:End Sub

Sub sw61_Spin:vpmTimer.PulseSw 61:PlaySoundAtVol "spinner", sw61, 1:End Sub 'left ramp spinner
Sub sw62_Hit:Me.TimerEnabled = 1:Controller.Switch(62) = 1:End Sub
Sub sw62_Timer:Me.TimerEnabled = 0:Controller.Switch(62) = 0:End Sub

 '****************************************
 ' SetLamp 0 is Off
 ' SetLamp 1 is On
 ' LampState(x) current state
 '****************************************

'Dim RefreshARlight
Dim LampState(200), FadingLevel(200), FadingState(200)
Dim FlashState(200), FlashLevel(200)
Dim FlashSpeedUp, FlashSpeedDown



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

GISetDefaultColorTimer.Interval = 1000
Sub GISetDefaultColorTimer_Timer  'If timer expires, no mode is running so set defaultGIcolor
  If GI_Light.State = 1 Then
    red = 0:green = 0:blue = 255
  End If
  GISetDefaultColorTimer.Enabled = 0
End Sub

Sub FlashInit
    Dim i
    For i = 0 to 200
        FlashState(i) = 0
        FlashLevel(i) = 0
    Next

    FlashSpeedUp = 50   ' fast speed when turning on the flasher
    FlashSpeedDown = 50 ' slow speed when turning off the flasher, gives a smooth fading
    AllFlashOff()
End Sub

Sub AllFlashOff
    Dim i
    For i = 0 to 200
        FlashState(i) = 0
    Next
End Sub

Sub Reflections_Timer()

End Sub


Sub UpdateLamps()


NFadeL  5, l5a16
NFadeL  6, l6a16
NFadeL  7, l7a16
NFadeL  8, l8a16
NFadeL  9, l9a16
NFadeL  10, l10a16
NFadeL  11, l11a16
NFadeL  12, l12a16
NFadeLm 13, l13a16
NFadeL  13, l13a16a
NFadeL  14, l14a16 : Light001.state=l14a16.state
NFadeL  15, l15a16 : Light002.state=l15a16.state
NFadeL  16, l16a16
NFadeLm 17, l17a16
NFadeL  17, l17a16a
NFadeLm 18, l18a16
NFadeL  18, l18a16a
NFadeLm 19, l19a16
NFadeL  19, l19a16a
NFadeLm 20, l20a16
NFadeL  20, l20a16a
NFadeLm 21, l21a16
NFadeL  21, l21a16a
NFadeLm 22, l22a16
NFadeL  22, l22a16a
NFadeLm 23, l23a16
NFadeL  23, l23a16a
NFadeLm 24, l24a16
NFadeL  24, l24a16a
NFadeL  25, l25a16
NFadeL  26, l26a16
NFadeL  27, l27a16
NFadeL  28, l28a16
NFadeL  29, l29a16
NFadeL  30, l30a16
NFadeL  31, l31a16
NFadeL  32, l32a16
NFadeL 33, l33a16
NFadeL 34, l34a16
NFadeL 35, l35a16
NFadeL 36, l36a16
NFadeL 37, l37a16
NFadeL 38, l38a16
NFadeL 39, l39a16
NFadeLm 41, l41a16
NFadeL  41, l41a16a
NFadeLm 42, l42a16
NFadeL  42, l42a16a
NFadeL  43, l43a16
NFadeL  44, l44a16
NFadeLm 45, l45a16
NFadeLm 45, l45f
NFadeL  46, l46a16
NFadeLm 47, l47a16
NFadeL  47, l47a16a
NFadeLm 48, l48a16
NFadeL  48, l48a16a
NFadeL  49, l49a16
NFadeL  50, l50a16
NFadeL  51, l51a16
NFadeL  52, l52a16
NFadeL  53, l53a16
NFadeL  54, l54a16
NFadeL  55, l55a16
NFadeL  56, l56a16
NFadeLm 57, l57a16
NFadeL  57, l57a16a
NFadeL 58, l58a16
NFadeL 59, l59a16
NFadeL 60, l60a16
NFadeL 61, l61a16
NFadeL 62, l62a16
NFadeL 63, l63a16
NFadeL 64, l64a16
NFadeL 65, l65a16
NFadeL 66, l66a16
NFadeL 67, l67a16
NFadeL 68, l68a16
NFadeL 69, l69a16
NFadeL 70, l70a16
NFadeL 71, l71a16
NFadeL 72, l72a16
NFadeL 73, l73a16

l46ar16.state = l46a16.state
l56ar16.state = l56a16.state
l69ar16.state = l69a16.state
l70ar16.state = l70a16.state
l71ar16.state = l71a16.state
l72ar16.state = l72a16.state
l73ar16.state = l73a16.state

    NFadeLm 40, l40a
    NFadeLm 40, l40b
    NFadeLm 40, l40a1
    NFadeL 40, l40b1

    NFadeLm 81, f26
        NFadeL 81, l81
If l81.state=1 Then l81.state=2
    NFadeLm 82, f26
        NFadeL 82, l82
If l82.state=1 Then l82.state=2
    NFadeLm 83, f26
        NFadeL 83, l83
If l83.state=1 Then l83.state=2
    NFadeLm 84, l84a
    NFadeL 84, GIWhite 'overall color
    'NFadeL 84, l84b
    NFadeLm 87, l87a
    NFadeL 87, GIRed 'overall color
    'NFadeL 87, l87b
    NFadeL 88, l88a
    'NFadeL 88, l88b
    NFadeL 89, l89a
    'NFadeL 89, l89b
    NFadeL 90, l90a
    'NFadeL 90, l90b
    NFadeL 91, l91
    NFadeL 92, l92a
    'NFadeL 92, l92b
    NFadeL 93, l93a
    'NFadeL 93, l93b
    NFadeL 94, l94a
    'NFadeL 94, l94b
    NFadeLm 85, l85a 'left flipper return
    NFadeLm 85, l85b 'left flipper return
    NFadeLm 85, l85e 'left flipper return
    'NFadeL 85, l85f 'left flipper return
    NFadeL 85, l85d 'left flipper return
    NFadeLm 95, l95e 'left slingshot
    NFadeLm 95, l95f 'left slingshot
    NFadeLm 95, l95a 'left slingshot
    'NFadeL 95, l95h 'left slingshot
    NFadeLm 86, l86a 'right flipper return
    NFadeLm 86, l86b 'right flipper return
    NFadeLm 86, l86c 'right flipper return
    'NFadeL 86, l86f 'right flipper return
    NFadeL 86, l86d 'right flipper return
    NFadeLm 96, l96e 'right slingshot
    NFadeLm 96, l96f 'right slingshot
    NFadeLm 96, l96a 'right slingshot
    'NFadeLm 96, l96b 'right slingshot
    NFadeLm 97, l97a 'upper lane guide k-i
    NFadeL 97, l97b 'upper lane guide k-i
    NFadeLm 98, l98a 'upper lane guide i-d
    NFadeL 98, l98b 'upper lane guide i-d
    NFadeL 99, l99a
    'NFadeL 99, l99b
    NFadeL 100, l100a
    'NFadeL 100, l100b
    NFadeL 101, l101a
    'NFadeL 101, l101b
    NFadeL 102, l102a
    'NFadeL 102, l102b
    NFadeL 103, l103a
    'NFadeL 103, l103b
    NFadeL 104, l104a
    'NFadeL 104, l104b
    NFadeL 105, l105a
    'NFadeL 105, l105b
    NFadeL 106, l106a
    'NFadeL 106, l106b
    NFadeL 107, l107a
    'NFadeL 107, l107b
    NFadeL 108, l108a
    'NFadeL 108, l108b
    NFadeL 109, l109
    NFadeL 110, l110
    NFadeL 111, l111a
    'NFadeL 111, l111b
    NFadeL 112, l112a
    'NFadeL 112, l112b
    NFadeL 113, l113
    NFadeL 114, l114
    NFadeL 115, l115a
    'NFadeL 115, l115b
    NFadeL 116, l116
    NFadeL 117, l117a
    'NFadeL 117, l117b
    NFadeL 118, l118a
    'NFadeL 118, l118b
    NFadeL 119, l119
    NFadeL 120, l120
    NFadeL 121, l121a 'left spinner flash
    Flash 121, l121b 'left spinner flash
    NFadeL 122, l122a 'right spinner flash
    Flash 122, l122b 'right spinner flash
    NFadeL 127, l127a
    'NFadeL 127, l127b
    'NFadeL 127, l127c
    'NFadeL 127, l127d
    NFadeLm 128, l128a
    'NFadeLm 128, l128b
    NFadeL 128, l128c
    'NFadeL 128, l128d
    NFadeLm 129, f23a 'hotel lex flash
    NFadeL 129, f23b 'hotel lex flash
    Flash 129, f23c 'hotel lex flash
    NFadeLm 130, f22a 'right ramp flash
    NFadeL 130, f22b 'right ramp flash
    Flash 130, f22c 'right ramp flash
    NFadeLm 131, f27a
    NFadeL 131, f27b

  If HighPerformance = 1 then
    FadePrim 140, Backwall, "backwall-kingpinON3", "backwall-kingpinON2", "backwall-kingpinON1","backwall-kingpin"
    FadePrim 132, Primitive_ramp1, "Ramp1RF_ON3", "Ramp1RF_ON2", "Ramp1RF_ON1","Ramp1map_OFF80"
    FadePrim 141, Primitive_ramp1decal, "Rampdecalredraw2ON3", "Rampdecalredraw2ON2", "Rampdecalredraw2ON1","Rampdecalredraw2"
    FadePrim 133, Primitive_ramp2, "Ramp2map_RFON3", "Ramp2map_RFON2", "Ramp2map_RFON1", "Ramp2map_OFF80"
    FadePrim 134, Primitive_rampcover, "PlasticCover_RFON3", "PlasticCover_RFON2", "PlasticCover_RFON1", "PlasticCoverMap"
'   FadePrim 135, Primitive_LWireReflect, "LwireReflect_ON3", "LwireReflect_ON2", "LwireReflect_ON1", "Off"
'   FadePrim 136, Primitive_Reflect2, "RwireReflect_ON3", "RwireReflect_ON2", "RwireReflect_ON1", "Off"
    FadePrim 137, HotelLEX, "HotelLex_RRFON3", "HotelLex_RRFON2", "HotelLex_RRFON1", "HotelLex"
    FadePrim 138, HotelLEX, "HotelLex_LRFON3", "HotelLex_LRFON2", "HotelLex_LRFON1", "HotelLex"
    FadePrim 139, HotelLEX, "HotelLex_2XRFON3", "HotelLex_2XRFON2", "HotelLex_2XRFON1", "HotelLex"
'   FadePrim 142, centerwirelower, "WireS_ON3", "WireS_ON2", "WireS_ON1", "Off"
'   FadePrim 143, p122b2, "wirebendR_ON3", "wirebendR_ON2", "wirebendR_ON1", "Off" 'right to left wire
'   FadePrim 144, p122b1, "wirebendL_ON3", "wirebendL_ON2", "wirebendL_ON1", "Off" 'left to right wire
  End If
End Sub


Sub SpinCityR_Timer()
  If HighPerformance = 1 then
    If l122b.opacity = 0 then p122a.image = "off":p122b.image = "off":end if
    If l122b.opacity >= 1 and l122b.opacity <= 33 then p122a.image = "Wire2_ON1":p122b.image = "wirebendY_ON1":end if
    If l122b.opacity >= 34 and l122b.opacity <= 66 then p122a.image = "Wire2_ON2":p122b.image = "wirebendY_ON2":end if
    If l122b.opacity >= 67 and l122b.opacity <= 100 then p122a.image = "Wire2_ON3":p122b.image = "wirebendY_ON3":end if
  End If
End Sub

Sub FadePrim(nr, pri, a, b, c, d)
    Select Case FadingLevel(nr)
        Case 2:pri.image = d:FadingLevel(nr) = 0
        Case 3:pri.image = c:FadingLevel(nr) = 2
        Case 4:pri.image = b:FadingLevel(nr) = 3
        Case 5:pri.image = a:FadingLevel(nr) = 1
    End Select
End Sub

Sub FadeL(nr, a, b)
    Select Case FadingLevel(nr)
        Case 2:b.state = 0:FadingLevel(nr) = 0
        Case 3:b.state = 1:FadingLevel(nr) = 2
        Case 4:a.state = 0:FadingLevel(nr) = 3
        Case 5:a.state = 1:FadingLevel(nr) = 1
    End Select
End Sub

Sub NFadeL(nr, a)
    Select Case FadingLevel(nr)
        Case 4:a.state = 0:FadingLevel(nr) = 0
        Case 5:a.State = 1:FadingLevel(nr) = 1
    End Select
End Sub

Sub NFadeLm(nr, a)
    Select Case FadingLevel(nr)
        Case 4:a.state = 0
        Case 5:a.State = 1
    End Select
End Sub

Sub Flash(nr, object)
    Select Case FlashState(nr)
        Case 0 'off
            FlashLevel(nr) = FlashLevel(nr) - FlashSpeedDown
            If FlashLevel(nr) < 0 Then
                FlashLevel(nr) = 0
                FlashState(nr) = -1 'completely off
            End if
            Object.opacity = FlashLevel(nr)
        Case 1 ' on
            FlashLevel(nr) = FlashLevel(nr) + FlashSpeedUp
            If FlashLevel(nr) > 100 Then
                FlashLevel(nr) = 100
                FlashState(nr) = -2 'completely on
            End if
            Object.opacity = FlashLevel(nr)
    End Select
End Sub

 Sub AllLampsOff():For x = 1 to 200:LampState(x) = 4:FadingLevel(x) = 4:Next:UpdateLamps:UpdateLamps:Updatelamps:End Sub


Sub SetLamp(nr, value)
    If value = 0 AND LampState(nr) = 0 Then Exit Sub
    If value = 1 AND LampState(nr) = 1 Then Exit Sub
    LampState(nr) = abs(value) + 4
FadingLevel(nr ) = abs(value) + 4: FadingState(nr ) = abs(value) + 4
End Sub

Sub SetFlash(nr, stat)
    FlashState(nr) = ABS(stat)
End Sub


Sub FlashAR(nr, ramp, a, b, c, r)                                                          'used for reflections when there is no off ramp
    Select Case FadingState(nr)
        Case 2:ramp.image = 0:r.State = ABS(r.state -1):FadingState(nr) = 0                'Off
        Case 3:ramp.image = c:r.State = ABS(r.state -1):FadingState(nr) = 2                'fading...
        Case 4:ramp.image = b:r.State = ABS(r.state -1):FadingState(nr) = 3                'fading...
        Case 5:ramp.image = a:ramp.alpha = 1:r.State = ABS(r.state -1):FadingState(nr) = 1 'ON
    End Select
End Sub


Sub FlashARm(nr, ramp, a, b, c, r)
    Select Case FadingState(nr)
        Case 2:ramp.alpha = 0:r.State = ABS(r.state -1)
        Case 3:ramp.image = c:r.State = ABS(r.state -1)
        Case 4:ramp.image = b:r.State = ABS(r.state -1)
        Case 5:ramp.image = a:ramp.alpha = 1:r.State = ABS(r.state -1)
    End Select
End Sub

Sub Flashm(nr, object) 'multiple flashers, it doesn't change the flashstate
    Select Case FlashState(nr)
        Case 0         'off
            Object.alpha = FlashLevel(nr)
        Case 1         ' on
            Object.alpha = FlashLevel(nr)
    End Select
End Sub

'SOUNDS
Dim leftdrop:leftdrop = 0
Sub leftdrop1_Hit:leftdrop = 1:End Sub
Sub leftdrop2_Hit
  If leftdrop = 1 then
    PlaySoundAtVol "drop_left", ActiveBall, 1
  End If
  StopSound "fx_metalrolling"
  leftdrop = 0
End Sub

Dim rightdrop:rightdrop = 0
Sub rightdrop1_Hit:rightdrop = 1:End Sub
Sub rightdrop2_Hit
  If rightdrop = 1 then
    PlaySoundAtVol "drop_Right", ActiveBall, 1
  End If
  StopSound "fx_metalrolling"
  rightdrop = 0
End Sub


Dim gistep
Dim gistep2

gistep = 255 / 8
gistep2 = 180 / 8

Sub RampsOff

End Sub

Sub RampsOn

End Sub

Sub FlippersOn

End Sub

Sub FlippersOff

End Sub

Sub FlippersRedOn

End Sub

'drop targets using flippers
Sub PrimT_Timer
  If sw25.isdropped = true then sw25f.rotatetoend
  if sw25.isdropped = false then sw25f.rotatetostart
  If sw26.isdropped = true then sw26f.rotatetoend
  if sw26.isdropped = false then sw26f.rotatetostart
  If sw27.isdropped = true then sw27f.rotatetoend
  if sw27.isdropped = false then sw27f.rotatetostart
  If sw28.isdropped = true then sw28f.rotatetoend
  if sw28.isdropped = false then sw28f.rotatetostart
  If sw29.isdropped = true then sw29f.rotatetoend
  if sw29.isdropped = false then sw29f.rotatetostart
  If sw30.isdropped = true then sw30f.rotatetoend
  if sw30.isdropped = false then sw30f.rotatetostart
  If sw31.isdropped = true then sw31f.rotatetoend
  if sw31.isdropped = false then sw31f.rotatetostart
  sw25p.transy = sw25f.currentangle
  sw26p.transy = sw26f.currentangle
  sw27p.transy = sw27f.currentangle
  sw28p.transy = sw28f.currentangle
  sw29p.transy = sw29f.currentangle
  sw30p.transy = sw30f.currentangle
  sw31p.transy = sw31f.currentangle
End Sub


Sub Table_exit()
  Controller.Pause = False
  Controller.Stop
End Sub

Dim ToggleMechSounds
Function SoundFX (sound)
  If cController = 4 and ToggleMechSounds = 0 Then
    SoundFX = ""
  Else
    SoundFX = sound
  End If
End Function

Sub ColorCheck_Timer
slotgi.State = l101a.State
slotgired.State = l102a.State
End Sub

Sub Pins_Hit (idx)
  PlaySound "pinhit_low", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioPan(ActiveBall)
End Sub

Sub Targets_Hit (idx)
  PlaySound ("Target_Hit_" & int(rnd(1)*9)+1) , 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioPan(ActiveBall)
End Sub

Sub Metals_Thin_Hit (idx)
' PlaySound "metalhit_thin", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
  PlaySound ("Metal_Touch_" & int(rnd(1)*13)+1) , 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioPan(ActiveBall)

End Sub

Sub Metals_Medium_Hit (idx)
' PlaySound "metalhit_medium", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
  PlaySound ("Metal_Touch_" & int(rnd(1)*13)+1) , 0, Vol(ActiveBall)+0.3, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioPan(ActiveBall)

End Sub

Sub Metals2_Hit (idx)
  PlaySound "metalhit2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioPan(ActiveBall)
End Sub

Sub Gates_Hit (idx)
  PlaySound "gate4", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioPan(ActiveBall)
End Sub

Sub Spinner_Spin
  PlaySoundAtVol "fx_spinner", Spinner, 1
End Sub

Sub Rubbers_Hit(idx)
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 20 then
'   PlaySound "fx_rubber2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
    PlaySound ("Rubber_strong_" & int(rnd(1)*9)+1) , 0, 1, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioPan(ActiveBall)
  End if
  If finalspeed >= 6 AND finalspeed <= 20 then
    PlaySound "Rubber_" & int(rnd(1)*9)+1 , 0, 1, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioPan(ActiveBall)
 '    RandomSoundRubber()
  End If
End Sub

Sub Posts_Hit(idx)
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 16 then
'   PlaySound "fx_rubber2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
    PlaySound ("Rubber_Strong_" & int(rnd(1)*9)+1) , 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioPan(ActiveBall)

  End if
  If finalspeed >= 6 AND finalspeed <= 16 then
    PlaySound ("Rubber_" & int(rnd(1)*9)+1) , 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioPan(ActiveBall)
'     RandomSoundRubber()
  End If
End Sub

Sub RandomSoundRubber()
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySound "rubber_hit_1", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioPan(ActiveBall)
    Case 2 : PlaySound "rubber_hit_2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioPan(ActiveBall)
    Case 3 : PlaySound "rubber_hit_3", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioPan(ActiveBall)
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
    Case 1 : PlaySound "flip_hit_1", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioPan(ActiveBall)
    Case 2 : PlaySound "flip_hit_2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioPan(ActiveBall)
    Case 3 : PlaySound "flip_hit_3", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioPan(ActiveBall)
  End Select
End Sub


Sub LRRail_Hit:PlaySound "fx_metalrolling", 0, 250, -0.88, 0, Pitch(ActiveBall), 1, 0, AudioPan(ActiveBall):End Sub

Sub RLRail_Hit:PlaySound "fx_metalrolling", 0, 250, 0.88, 0, Pitch(ActiveBall), 1, 0, AudioPan(ActiveBall):End Sub

' *******************************************************************************************************
' Positional Sound Playback Functions by DJRobX, Rothbauerw, Thalamus and Herweh
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

' set position as table object and Vol + RndPitch manually

Sub PlaySoundAtVolPitch(sound, tableobj, Vol, RndPitch)
  PlaySound sound, 1, Vol, AudioPan(tableobj), RndPitch, 0, 0, 1, AudioFade(tableobj)
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

Sub PlaySoundAtBallAbsVol(sound, VolMult)
  PlaySound sound, 0, VolMult, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 1, AudioFade(ActiveBall)
End Sub

' requires rampbump1 to 7 in Sound Manager

Sub RandomBump(voladj, freq)
  Dim BumpSnd:BumpSnd= "rampbump" & CStr(Int(Rnd*7)+1)
  PlaySound BumpSnd, 0, Vol(ActiveBall)*voladj, AudioPan(ActiveBall), 0, freq, 0, 1, AudioFade(ActiveBall)
End Sub

' set position as bumperX and Vol manually. Allows rapid repetition/overlaying sound

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
'            Supporting Ball, Sound Functions and Math
'*********************************************************************

Function RndNum(min, max)
  RndNum = Int(Rnd() * (max-min + 1) ) + min ' Sets a random number between min and max
End Function

Const Pi = 3.1415927

Function dSin(degrees)
  dsin = sin(degrees * Pi/180)
End Function

Function dCos(degrees)
  dcos = cos(degrees * Pi/180)
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
  Vol = Csng(BallVel(ball) ^2 / 2000)
End Function

Function VolMulti(ball,Multiplier) ' Calculates the Volume of the sound based on the ball speed
  VolMulti = Csng(BallVel(ball) ^2 / 150 ) * Multiplier
End Function

Function DVolMulti(ball,Multiplier) ' Calculates the Volume of the sound based on the ball speed
  DVolMulti = Csng(BallVel(ball) ^2 / 150 ) * Multiplier
  debug.print DVolMulti
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


''*****************************************
''      JP's VP10 Rolling Sounds
''*****************************************

Const tnob = 8 ' total number of balls
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
        If BallVel(BOT(b) ) > 1 AND BOT(b).z < 30 AND BOT(b).z > 0 Then
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
' PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 2000, Pan(ball1), 0, Pitch(ball1), 0, 0
  PlaySound ("Ball_Collide_" &  int(rnd(1)*7)+1 ), 0, Csng(velocity) ^2 / 2000, Pan(ball1), 0, Pitch(ball1), 0, 0
End Sub



Const ReflipAngle = 20
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
    x.TimeDelay = 80
  Next

  'rf.report "Velocity"
  AddPt "Polarity", 0, 0, 0
  AddPt "Polarity", 1, 0.05, -3.7
  AddPt "Polarity", 2, 0.33, -3.7
  AddPt "Polarity", 3, 0.37, -3.7
  AddPt "Polarity", 4, 0.41, -3.7
  AddPt "Polarity", 5, 0.45, -3.7
  AddPt "Polarity", 6, 0.576,-3.7
  AddPt "Polarity", 7, 0.66, -2.3
  AddPt "Polarity", 8, 0.743, -1.5
  AddPt "Polarity", 9, 0.81, -1
  AddPt "Polarity", 10, 0.88, 0

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
End Sub

Sub AddPt(aStr, idx, aX, aY)  'debugger wrapper for adjusting flipper script in-game
  dim a : a = Array(LF, RF)
  dim x : for each x in a
    x.addpoint aStr, idx, aX, aY
  Next
End Sub

'Trigger Hit - .AddBall activeball
'Trigger UnHit - .PolarityCorrect activeball
dim LFFlipBall, RFFlipBall : LFFlipBall = 0 : RFFlipBall = 0


Sub TriggerLF_Hit() : LF.Addball activeball : LFFlipBall = 1 : End Sub
Sub TriggerLF_UnHit() : LF.PolarityCorrect activeball : LFFlipBall = 0 : End Sub
Sub TriggerRF_Hit() : RF.Addball activeball : RFFlipBall = 1 : End Sub
Sub TriggerRF_UnHit() : RF.PolarityCorrect activeball : RFFlipBall = 0 : End Sub

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
  private Flipper, FlipperStart,FlipperEnd, FlipperEndY, LR, PartialFlipCoef
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
  Public Property Let EndPoint(aInput) : FlipperEnd = aInput.x: FlipperEndY = aInput.y: End Property
  Public Property Get EndPoint : EndPoint = FlipperEnd : End Property
  Public Property Get EndPointY: EndPointY = FlipperEndY : End Property


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
      End If
    Next
    PartialFlipCoef = ((Flipper.StartAngle - Flipper.CurrentAngle) / (Flipper.StartAngle - Flipper.EndAngle))
    PartialFlipCoef = abs(PartialFlipCoef-1)
  End Sub
  Private Function FlipperOn() : if gameTime < FlipAt+TimeDelay then FlipperOn = True : End If : End Function 'Timer shutoff for polaritycorrect

  Public Sub PolarityCorrect(aBall)
    if FlipperOn() then
      dim tmp, BallPos, x, IDX, Ycoef : Ycoef = 1

      'y safety Exit
      if aBall.VelY > -8 then 'ball going down
        RemoveBall aBall
        exit Sub
      end if

      'Find balldata. BallPos = % on Flipper
      for x = 0 to uBound(Balls)
        if aBall.id = BallData(x).id AND not isempty(BallData(x).id) then
          idx = x
          BallPos = PSlope(BallData(x).x, FlipperStart, 0, FlipperEnd, 1)
          if ballpos > 0.65 then  Ycoef = LinearEnvelope(BallData(x).Y, YcoefIn, YcoefOut)        'find safety coefficient 'ycoef' data
        end if
      Next

      If BallPos = 0 Then 'no ball data meaning the ball is entering and exiting pretty close to the same position, use current values.
        BallPos = PSlope(aBall.x, FlipperStart, 0, FlipperEnd, 1)
        if ballpos > 0.65 then  Ycoef = LinearEnvelope(aBall.Y, YcoefIn, YcoefOut)            'find safety coefficient 'ycoef' data
      End If

      'Velocity correction
      if not IsEmpty(VelocityIn(0) ) then
        Dim VelCoef
   :      VelCoef = LinearEnvelope(BallPos, VelocityIn, VelocityOut)

        if partialflipcoef < 1 then VelCoef = PSlope(partialflipcoef, 0, 1, 1, VelCoef)

        if Enabled then aBall.Velx = aBall.Velx*VelCoef
        if Enabled then aBall.Vely = aBall.Vely*VelCoef
      End If

      'Polarity Correction (optional now)
      if not IsEmpty(PolarityIn(0) ) then
        If StartPoint > EndPoint then LR = -1 'Reverse polarity if left flipper
        dim AddX : AddX = LinearEnvelope(BallPos, PolarityIn, PolarityOut) * LR

        if Enabled then aBall.VelX = aBall.VelX + 1 * (AddX*ycoef*PartialFlipcoef)
        'debug.print BallPos & " " & AddX & " " & Ycoef & " "& PartialFlipcoef & " "& VelCoef
        'playsound "fx_knocker"
      End If
    End If
    RemoveBall aBall
  End Sub
End Class


'================================
'Helper Functions


Sub ShuffleArray(ByRef aArray, byVal offset) 'shuffle 1d array
  dim x, aCount : aCount = 0
  redim ee(uBound(aArray) )
  for x = 0 to uBound(aArray) 'Shuffle objects in a temp array
    if not IsEmpty(aArray(x) ) Then
      if IsObject(aArray(x)) then
        Set ee(aCount) = aArray(x)
      Else
        ee(aCount) = aArray(x)
      End If
      aCount = aCount + 1
    End If
  Next
  if offset < 0 then offset = 0
  redim aArray(aCount-1+offset) 'Resize original array
  for x = 0 to aCount-1   'set objects back into original array
    if IsObject(ee(x)) then
      Set aArray(x) = ee(x)
    Else
      aArray(x) = ee(x)
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

Function Distance(ax,ay,bx,by)
  Distance = SQR((ax - bx)^2 + (ay - by)^2)
End Function

Function DistancePL(px,py,ax,ay,bx,by) ' Distance between a point and a line where point is px,py
  DistancePL = ABS((by - ay)*px - (bx - ax) * py + bx*ay - by*ax)/Distance(ax,ay,bx,by)
End Function

'******************** Slacky Flips **************************'

Const FlipY = 1858           'Default flip Y position
Const Dislocation = 2        'Amount of minimum slack, affect to recovery speed too
Const SlackHitLimit = 6    '6-16    Lower value will make slack happen with lower collision force
Const HitForceDivider = 10    '10-100    Lower value will add slack. Set to roughly half of the max hit force(parm) you see when playing
Const MaxSlack = 8
Dim SlackAmount

Sub CheckFlipperSlack(ball, Flipper, parm)
    If slackyFlips = 1 Then
        SlackAmount = Dislocation + parm / HitForceDivider
        if SlackAmount > MaxSlack then SlackAmount = MaxSlack
        if parm > SlackHitLimit Then
            Flipper.Y = FlipY + SlackAmount
           ' Logo.Y = FlipY + SlackAmount
            slacktimer.interval = 10
            slacktimer.Enabled = 1
            'debug.print parm & " parm and SlackAmount " & SlackAmount
        end If
    end If
end Sub

Sub slacktimer_timer()
    'msgbox "location change" & RightFlipper.Y & " and " & LeftFlipper.Y
    if RightFlipper.Y > FlipY + 2 Then
        RightFlipper.Y = RightFlipper.Y - 2
      '  RFLogo.Y = RFLogo.Y - 2
    else
        RightFlipper.Y = FlipY
       ' RFLogo.Y = FlipY
    End If
    if LeftFlipper.Y > FlipY + 2 Then
        LeftFlipper.Y = LeftFlipper.Y - 2
       ' LFLogo.Y = LFLogo.Y - 2
    else
        LeftFlipper.Y = FlipY
      '  LFLogo.Y = FlipY
    End If

    if LeftFlipper.Y = FlipY and RightFlipper.Y = FlipY Then slacktimer.Enabled = 0
end Sub

'******************************************************
'     FLIPPER TRICKS
'******************************************************

RightFlipper.timerinterval=1
Rightflipper.timerenabled=True

sub RightFlipper_timer()
  FlipperTricks LeftFlipper, LFPress, LFCount, LFEndAngle, LFState
  FlipperTricks RightFlipper, RFPress, RFCount, RFEndAngle, RFState
  If FlipSolNudge = 1 Then
    FlipperNudge RightFlipper, RFEndAngle, RFEOSNudge, LeftFlipper, LFEndAngle
    FlipperNudge LeftFlipper, LFEndAngle, LFEOSNudge,  RightFlipper, RFEndAngle
  End If
end sub

Dim LFEOSNudge, RFEOSNudge

Sub FlipperNudge(Flipper1, Endangle1, EOSNudge1, Flipper2, EndAngle2)
  Dim BOT, b

  If Flipper1.currentangle = Endangle1 and EOSNudge1 <> 1 Then
    EOSNudge1 = 1
    If Flipper2.currentangle = EndAngle2 Then
      BOT = GetBalls
        For b = 0 to Ubound(BOT)
        If FlipperTrigger(BOT(b).x, BOT(b).y, Flipper2) Then
          BOT(b).velx = BOT(b).velx / 1.7
          BOT(b).vely = BOT(b).vely - 1
        end If
      Next
    End If
  Else
    If Flipper1.currentangle <> EndAngle1 then EOSNudge1 = 0
  End If
End Sub

'*************************************************
' Check ball distance from Flipper for Rem
'*************************************************

Function Distance(ax,ay,bx,by)
  Distance = SQR((ax - bx)^2 + (ay - by)^2)
End Function

Function DistancePL(px,py,ax,ay,bx,by) ' Distance between a point and a line where point is px,py
  DistancePL = ABS((by - ay)*px - (bx - ax) * py + bx*ay - by*ax)/Distance(ax,ay,bx,by)
End Function

Function Radians(Degrees)
  Radians = Degrees * PI /180
End Function

Function AnglePP(ax,ay,bx,by)
  AnglePP = Atn2((by - ay),(bx - ax))*180/PI
End Function

Function Atn2(dy, dx)
  If dx > 0 Then
    Atn2 = Atn(dy / dx)
  ElseIf dx < 0 Then
    Atn2 = Sgn(dy) * (PI - Atn(Abs(dy / dx)))
  ElseIf dy = 0 Then
    Atn2 = 0
  Else
    Atn2 = Sgn(dy) * PI / 2
  End If
End Function

Function DistanceFromFlipper(ballx, bally, Flipper)
  DistanceFromFlipper = DistancePL(ballx, bally, Flipper.x, Flipper.y, Cos(Radians(Flipper.currentangle+90))+Flipper.x, Sin(Radians(Flipper.currentangle+90))+Flipper.y)
End Function

Function FlipperTrigger(ballx, bally, Flipper)
  Dim DiffAngle
  DiffAngle  = ABS(Flipper.currentangle - AnglePP(Flipper.x, Flipper.y, ballx, bally) - 90)
  If DiffAngle > 180 Then DiffAngle = DiffAngle - 360

  If DistanceFromFlipper(ballx,bally,Flipper) < 48 and DiffAngle <= 90 and Distance(ballx,bally,Flipper.x,Flipper.y) < Flipper.Length Then
    FlipperTrigger = True
  Else
    FlipperTrigger = False
  End If
End Function

'*************************************************
' End - Check ball distance from Flipper for Rem
'*************************************************



dim LFPress, RFPress, LFCount, RFCount
dim LFState, RFState
dim EOST, EOSA,Frampup, FElasticity,FReturn
dim RFEndAngle, LFEndAngle

EOST = leftflipper.eostorque
EOSA = leftflipper.eostorqueangle
Frampup = LeftFlipper.rampup
FElasticity = LeftFlipper.elasticity
FReturn = LeftFlipper.return
Const EOSTnew = 1
Const EOSAnew = 0.8
Const EOSRampup = 0
Dim SOSRampup
SOSRampup = 2.5
'Const LiveCatch = 16
Const LiveElasticity = 0.45
Const SOSEM = 0.815
Const EOSReturn = 0.025

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
  Flipper.eostorque = EOST*EOSReturn/FReturn


  If Abs(Flipper.currentangle) <= Abs(Flipper.endangle) + 0.1 Then
    Dim BOT, b
    BOT = GetBalls

    For b = 0 to UBound(BOT)
      If Distance(BOT(b).x, BOT(b).y, Flipper.x, Flipper.y) < 55 Then 'check for cradle
        If BOT(b).vely >= -0.4 Then BOT(b).vely = -0.4
      End If
    Next
  End If
End Sub

Sub FlipperTricks (Flipper, FlipperPress, FCount, FEndAngle, FState)
  Dim Dir
  Dir = Flipper.startangle/Abs(Flipper.startangle)  '-1 for Right Flipper

  If Abs(Flipper.currentangle) > Abs(Flipper.startangle) - 0.05 Then
    If FState <> 1 Then
      Flipper.rampup = SOSRampup
      Flipper.endangle = FEndAngle - 3*Dir
      Flipper.Elasticity = FElasticity * SOSEM
      FCount = 0
      FState = 1
    End If
  ElseIf Abs(Flipper.currentangle) <= Abs(Flipper.endangle) and FlipperPress = 1 then
    if FCount = 0 Then FCount = GameTime

    If FState <> 2 Then
      Flipper.eostorqueangle = EOSAnew
      Flipper.eostorque = EOSTnew
      Flipper.rampup = EOSRampup
      Flipper.endangle = FEndAngle
      FState = 2
    End If
  Elseif Abs(Flipper.currentangle) > Abs(Flipper.endangle) + 0.01 and FlipperPress = 1 Then
    If FState <> 3 Then
      Flipper.eostorque = EOST
      Flipper.eostorqueangle = EOSA
      Flipper.rampup = Frampup
      Flipper.Elasticity = FElasticity
      FState = 3
    End If

  End If
End Sub

Const LiveDistanceMin = 30  'minimum distance in vp units from flipper base live catch dampening will occur
Const LiveDistanceMax = 114  'maximum distance in vp units from flipper base live catch dampening will occur (tip protection)

Sub CheckLiveCatch(ball, Flipper, FCount, parm) 'Experimental new live catch
  Dim Dir
  Dir = Flipper.startangle/Abs(Flipper.startangle)    '-1 for Right Flipper
  Dim LiveCatchBounce                             'If live catch is not perfect, it won't freeze ball totally
  Dim CatchTime : CatchTime = GameTime - FCount

  if CatchTime <= LiveCatch and parm > 6 and ABS(Flipper.x - ball.x) > LiveDistanceMin and ABS(Flipper.x - ball.x) < LiveDistanceMax Then
    if CatchTime <= LiveCatch*0.5 Then            'Perfect catch only when catch time happens in the beginning of the window
      LiveCatchBounce = 0
    else
      LiveCatchBounce = Abs((LiveCatch/2) - CatchTime)  'Partial catch when catch happens a bit late
    end If

    If LiveCatchBounce = 0 and ball.velx * Dir > 0 Then ball.velx = 0
    ball.vely = LiveCatchBounce * (16 / LiveCatch) ' Multiplier for inaccuracy bounce
    ball.angmomx= 0
    ball.angmomy= 0
    ball.angmomz= 0
    debug.print "Live catch! Time: " & CatchTime & " Bounce: " & LiveCatchBounce & " parm: " & parm
  End If

End Sub

'*****************************************************************************************************
'*******************************************************************************************************
'END nFOZZY FLIPPERS'
'****************************************************************************
'nFozzy PHYSICS DAMPENERS

'These are data mined bounce curves,
'dialed in with the in-game elasticity as much as possible to prevent angle / spin issues.
'Requires tracking ballspeed to calculate COR

Sub dPosts_Hit(idx)
  RubbersD.dampen Activeball
  PlaySound ("Rubber_" & int(rnd(1)*9)+1 ), 0, 1, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioPan(ActiveBall)
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

    If cor.BallVel(aball.id) = Not 0 Then
      RealCOR = BallSpeed(aBall) / cor.ballvel(aBall.id)
      coef = desiredcor / realcor
      if debugOn then str = name & " in vel:" & round(cor.ballvel(aBall.id),2 ) & vbnewline & "desired cor: " & round(desiredcor,4) & vbnewline & _
      "actual cor: " & round(realCOR,4) & vbnewline & "ballspeed coef: " & round(coef, 3) & vbnewline
      if Print then debug.print Round(cor.ballvel(aBall.id),2) & ", " & round(desiredcor,3)
      aBall.velx = aBall.velx * coef : aBall.vely = aBall.vely * coef
      'playsound "fx_knocker"
      if debugOn then TBPout.text = str
    End If
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

'******************************************************
'   TRACK ALL BALL VELOCITIES
'     FOR RUBBER DAMPENER AND DROP TARGETS
'******************************************************

dim cor : set cor = New CoRTracker

Class CoRTracker
  public ballvel, ballvelx, ballvely

  Private Sub Class_Initialize : redim ballvel(0) : redim ballvelx(0): redim ballvely(0) : End Sub

  Public Sub Update() 'tracks in-ball-velocity
    dim str, b, AllBalls, highestID : allBalls = getballs

    for each b in allballs
      if b.id >= HighestID then highestID = b.id
    Next

    if uBound(ballvel) < highestID then redim ballvel(highestID)  'set bounds
    if uBound(ballvelx) < highestID then redim ballvelx(highestID)  'set bounds
    if uBound(ballvely) < highestID then redim ballvely(highestID)  'set bounds

    for each b in allballs
      ballvel(b.id) = BallSpeed(b)
      ballvelx(b.id) = b.velx
      ballvely(b.id) = b.vely
    Next
  End Sub
End Class

Sub RDampen_Timer()
  Cor.Update
End Sub

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

'*******************************************************
' End nFozzy Dampening'
'******************************************************

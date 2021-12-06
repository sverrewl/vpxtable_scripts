Option Explicit
Randomize

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Const cGameName="jb_10r",UseSolenoids=2,UseLamps=0,UseGI=1,SSolenoidOn="SolOn",SSolenoidOff="SolOff", SCoin="coin"

' UseVPMModSol   When True this allows the ROM to control the intensity level of modulated solenoids
' instead of just on/off.
Const UseVPMModSol=1

LoadVPM "01560000", "WPC.VBS", 3.50  'LoadVPM "01570000", "wpc.vbs", 3.26
Dim DesktopMode: DesktopMode = Table1.ShowDT

If DesktopMode = True Then 'Show Desktop components
Ramp16.visible=1
Ramp15.visible=1
Else
Ramp16.visible=0
Ramp15.visible=0
End if

dim ballsize,ballmass

ballsize=1
ballmass=1.2

Dim NullFader : set NullFader = new NullFadingObject
Dim FadeLights : Set FadeLights = New LampFader
Dim FlasherInterval: FlasherInterval = 17


'**********************************************************************************************************

'Initiate Table
'**********************************************************************************************************

Dim bsTrough, bsSaucer, bsREye, bsLEye, dtDTBank, mVisor

Sub Table1_Init
   SetLocale(1033)
   solRampDwn(True)

  vpmInit Me
  On Error Resume Next
    With Controller
    .GameName = cGameName
    If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
    .SplashInfoLine = "Pinbot"&chr(13)&"bord"
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
    .Switch(22) = 1     'close coin door
    .Switch(24) = 1     'always closed
     End With

    PinMAMETimer.Interval = PinMAMEInterval
    PinMAMETimer.Enabled = 1

  vpmNudge.TiltSwitch=14
    vpmNudge.Sensitivity=1
    vpmNudge.TiltObj=Array(sw61_Bumper1,sw62_Bumper2,sw63_Bumper3,LeftSlingshot,RightSlingshot)

           Set bsTrough = New cvpmBallStack
         bstrough.InitSw 0, 32, 33, 34, 35, 0, 0, 0
             bsTrough.InitKick BallRelease, 160, 8
               bsTrough.InitExitSnd SoundFX("ballrelease",DOFContactors), SoundFX("Solenoid",DOFContactors)
               bsTrough.Balls = 4

          Set bsSaucer = New cvpmBallStack
              bsSaucer.InitSaucer sw46,46, 155, 12
            bsSaucer.KickForceVar = 2
            bsSaucer.KickAngleVar = 2
              bsSaucer.InitExitSnd SoundFX("Popper_ball",DOFContactors), SoundFX("Solenoid",DOFContactors)

          Set bsLEye = New cvpmBallStack
              bsLEye.InitSaucer sw47,47, 165, 12
            bsLEye.KickForceVar = 2
            bsLEye.KickAngleVar = 2
              bsLEye.InitExitSnd SoundFX("Popper_ball",DOFContactors), SoundFX("Solenoid",DOFContactors)

          Set bsREye = New cvpmBallStack
              bsREye.InitSaucer sw48,48, 185, 12
            bsREye.KickForceVar = 2
            bsREye.KickAngleVar = 2
              bsREye.InitExitSnd SoundFX("Popper_ball",DOFContactors), SoundFX("Solenoid",DOFContactors)

           set dtDTBank = new cvpmdroptarget
               dtDTBank.InitDrop Array(sw18, sw17, sw16), Array(18, 17, 16)
               dtDTBank.Initsnd SoundFX("DTDrop",DOFContactors),SoundFX("DTReset",DOFContactors)

  Set mVisor = New cvpmMech
  With mVisor
   .MType = vpmMechOneSol + vpmMechReverse + vpmMechLinear
         .Sol1 = 28
         .Length = 120
         .Steps = 58
         .AddSw 115, 0, 0
         .AddSw 117, 58, 58
         .Callback = GetRef("UpdateVisor")
         .Start
     End With

  Flasherset15 0
  Flasherset16 0
  Flasherset17 0
  Flasherset18 0
  Flasherset19 0
  Flasherset20 0
  Flasherset21 0
  Flasherset22 0
  Flasherset23 0
  Flasherset24 0
  Flasherset25 0
  Flasherset26 0
  Flasherset27 0

  FlasherInit.enabled = True

End Sub

Sub Table1_Exit(): Controller.Stop:End Sub

Dim FICount:FICount = 0

Sub FlasherInit_Timer()
  FICount = FICount + 1

  Select Case FICount
    Case 25:
      UpdateGi 0,8
      UpdateGi 1,8
      UpdateGi 2,8
      UpdateGi 3,8
      SetFlashersOdd 255
    Case 30:
      SetFlashersEven 255
    Case 35:
      SetFlashersOdd 0
    Case 40:
      SetFlashersEven 0
      UpdateGi 0,0
      UpdateGi 1,0
      UpdateGi 2,0
      UpdateGi 3,0
    Case 45:
      SetFlashersOdd 255
    Case 50:
      SetFlashersEven 255
    Case 55:
      SetFlashersOdd 0
    Case 60:
      SetFlashersEven 0
      UpdateGi 0,8
      UpdateGi 1,8
      UpdateGi 2,8
      UpdateGi 3,8
      solRampDwn 1
    Case 65: me.enabled = False
  End Select
End Sub

Sub SetFlashersOdd(value)
  Flasherset15 value
  Flasherset17 value
  Flasherset19 value
  Flasherset21 value
  Flasherset23 value
  Flasherset25 value
  Flasherset27 value
End Sub

Sub SetFlashersEven(value)
  Flasherset16 value
  Flasherset18 value
  Flasherset20 value
  Flasherset22 value
  Flasherset24 value
  Flasherset26 value
End Sub

'**********************************************************************************************************
'Plunger code
'**********************************************************************************************************

Sub Table1_KeyDown(ByVal KeyCode)
  If KeyDownHandler(keycode) Then Exit Sub
  If keycode = PlungerKey Then Plunger.Pullback:playsoundat"plungerpull", plunger
  if KeyCode = LeftTiltKey Then Nudge 90, 4
  if KeyCode = RightTiltKey Then Nudge 270, 4
  if KeyCode = CenterTiltKey Then Nudge 0, 4
'    if keycode = keyFront Then Controller.Switch(23) = 1
    If KeyCode = RightMagnaSave Then Controller.Switch(23) = 1
    ' Manual Ball Control
  If keycode = 46 Then          ' C Key
    If EnableBallControl = 1 Then
      EnableBallControl = 0
    Else
      EnableBallControl = 1
    End If
  End If
    If EnableBallControl = 1 Then
    If keycode = 48 Then        ' B Key
      If BCboost = 1 Then
        BCboost = BCboostmulti
      Else
        BCboost = 1
      End If
    End If
    If keycode = 203 Then BCleft = 1  ' Left Arrow
    If keycode = 200 Then BCup = 1    ' Up Arrow
    If keycode = 208 Then BCdown = 1  ' Down Arrow
    If keycode = 205 Then BCright = 1 ' Right Arrow
  End If
End Sub

Sub Table1_KeyUp(ByVal KeyCode)
  If KeyUpHandler(keycode) Then Exit Sub
  If keycode = PlungerKey Then Plunger.Fire:PlaySoundat"plunger", plunger
'    If keycode = keyFront Then Controller.Switch(23) = 0
    If KeyCode = RightMagnaSave Then Controller.Switch(23) = 0
    'Manual Ball Control
  If EnableBallControl = 1 Then
    If keycode = 203 Then BCleft = 0  ' Left Arrow
    If keycode = 200 Then BCup = 0    ' Up Arrow
    If keycode = 208 Then BCdown = 0  ' Down Arrow
    If keycode = 205 Then BCright = 0 ' Right Arrow
  End If
End Sub


'*************************************************************
'Solenoid Call backs
'**********************************************************************************************************
SolCallback(1) = "bsTrough.SolOut"
'2 not used
SolCallback(3) = "bsSaucer.SolOut"
SolCallback(4) ="dtDTBank.SolDropUp" 'Drop Targets
SolCallback(5) = "bsREye.SolOut"
SolCallback(6) = "solRampUp"
SolCallback(7) =  "vpmSolSound SoundFX(""Knocker"",DOFKnocker),"
SolCallback(8) = "bsLEye.SolOut"
'9 left slingshot
'10 right slingshot
'11 lower bumper
'12 left bumper
'13 upper bumper
SolCallback(14) = "solRampDwn"

SolModCallback(15) = "Flasherset15" 'RightEye Flash
SolModCallback(16) = "Flasherset16" 'LeftEye Flash
SolModCallback(17) = "Flasherset17" 'Center Visor Flash
SolModCallback(18) = "Flasherset18" 'Pinbot face flasher
SolModCallback(19) = "Flasherset19" 'Bumper Flasher
SolModCallback(20) = "Flasherset20" 'Flasher Lower Left
SolModCallback(21) = "Flasherset21" 'Flashers Center Left
SolModCallback(22) = "Flasherset22" 'Flasher Lower Right
SolModCallback(23) = "Flasherset23" 'Flashers 1 (L)
SolModCallback(24) = "Flasherset24" 'Flashers 2
SolModCallback(25) = "Flasherset25" 'Flashers 3
SolModCallback(26) = "Flasherset26" 'Flashers 4
SolModCallback(27) = "Flasherset27" 'Flashers 5 (R)

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
     If Enabled Then
         PlaySound SoundFX("fx_FlipperupL",DOFFlippers):LeftFlipper.RotateToEnd
     Else
         PlaySound SoundFX("fx_Flipperdown",DOFFlippers):LeftFlipper.RotateToStart
     End If
  End Sub

Sub SolRFlipper(Enabled)
     If Enabled Then
         PlaySound SoundFX("fx_FlipperupR",DOFFlippers):RightFlipper.RotateToEnd
     Else
         PlaySound SoundFX("fx_Flipperdown",DOFFlippers):RightFlipper.RotateToStart
     End If
End Sub

Sub DropTimer_Timer
  If GI1isoff=true then dropplate1.visible=0
  If GI1isoff=true then dropplate2.visible=0
  If GI1isoff=true then dropplate3.visible=0
  If GI1isoff=false and sw16.isdropped=1 then dropplate1.visible=0
  If GI1isoff=false and sw16.isdropped=0 then dropplate1.visible=1
  If GI1isoff=false and sw17.isdropped=1 then dropplate2.visible=0
  If GI1isoff=false and sw17.isdropped=0 then dropplate2.visible=1
  If GI1isoff=false and sw18.isdropped=1 then dropplate3.visible=0
  If GI1isoff=false and sw18.isdropped=0 then dropplate3.visible=1
End Sub
'
dim Gi0IsOff,Gi1IsOff,Gi2IsOff,Gi3IsOff


'*****************************************
'         Modulated Flashers
'*****************************************

Dim FlashFadeUp, FlashFadeDown

FlashFadeUp = 80
FlashFadeDown = 60

Dim FlashLevel15, FlashLevel16, FlashLevel17, FlashLevel18, FlashLevel19, FlashLevel20, FlashLevel21, FlashLevel22
Dim FlashLevel23, FlashLevel24, FlashLevel25, FlashLevel26,FlashLevel27

Dim Flash2Level15, Flash2Level16, Flash2Level17, Flash2Level18, Flash2Level19, Flash2Level20, Flash2Level21, Flash2Level22
Dim Flash2Level23, Flash2Level24, Flash2Level25, Flash2Level26,Flash2Level27

Dim FlashLevel15dir, FlashLevel16dir, FlashLevel17dir, FlashLevel18dir, FlashLevel19dir, FlashLevel20dir, FlashLevel21dir, FlashLevel22dir
Dim FlashLevel23dir, FlashLevel24dir, FlashLevel25dir, FlashLevel26dir,FlashLevel27dir

Sub FlasherClick(oldvalue, newvalue) : If oldvalue <= 0 and newvalue > 0.2 Then PlaySound "fx_relay",0,0.1: End If : End Sub

' ********* Right Eye Flasher **********

' FadeLights.obj(101) = array(flash101) 'right eye flash

Sub Flasherset15(value)
  if value < FlashLevel15 Then
    FlashLevel15dir =  -1
  elseif value > FlashLevel15 Then
    FlashLevel15dir = 1
  end If

  FlashLevel15 = value
  Flash101.Timerenabled = true

End Sub

Sub Flash101_timer()
  me.timerinterval = FlasherInterval

  If FlashLevel15dir > 0 Then
    If Flash2Level15 < FlashLevel15 Then
      Flash2Level15 = Flash2Level15 + FlashFadeUp
    Else
      Flash2Level15 = FlashLevel15
      me.timerenabled = False
    End If
  Else
    If Flash2Level15 > FlashLevel15 Then
      Flash2Level15 = Flash2Level15 - FlashFadeDown
    Else
      Flash2Level15 = FlashLevel15
      me.timerenabled = False
    End If
  End If

  if Flash2Level15 >  0 Then
    flash101.visible = 1
    flash101.IntensityScale = Flash2Level15 / 255
  else
    flash101.visible = 0
  end if

  if Flash2Level15 > 0 Then
    if Flash2Level16 > 0 Then
      mask_prim.image="maskon"
    else
      mask_prim.image="maskrighton" '101 is on, 104 is off
    end if
  else
    if Flash2Level16 > 0 Then
      mask_prim.image="masklefton"
    else
      mask_prim.image="maskoff"
    end if
  end if
End Sub

' ********* Left Eye Flasher **********

' FadeLights.obj(104) = array(flash104) 'left eye flash
Sub Flasherset16(value)
  if value < FlashLevel16 Then
    FlashLevel16dir =  -1
  elseif value > FlashLevel16 Then
    FlashLevel16dir = 1
  end If

  FlashLevel16 = value
  Flash104.Timerenabled = true

End Sub

Sub Flash104_timer()
  me.timerinterval = FlasherInterval

  If FlashLevel16dir > 0 Then
    If Flash2Level16 < FlashLevel16 Then
      Flash2Level16 = Flash2Level16 + FlashFadeUp
    Else
      Flash2Level16 = FlashLevel16
      me.timerenabled = False
    End If
  Else
    If Flash2Level16 > FlashLevel16 Then
      Flash2Level16 = Flash2Level16 - FlashFadeDown
    Else
      Flash2Level16 = FlashLevel16
      me.timerenabled = False
    End If
  End If

  if Flash2Level16 >  0 Then
    flash104.visible = 1
    flash104.IntensityScale = Flash2Level16 / 255
  else
    flash104.visible = 0
  end if

  if Flash2Level16 > 0 Then
    if Flash2Level16 > 0 Then
      mask_prim.image="maskon"
    else
      mask_prim.image="masklefton" '104 is on, 101 is off
    end if
  else
    if Flash2Level15 > 0 Then
      mask_prim.image="maskrighton"
    else
      mask_prim.image="maskoff"
    end if
  end if
End Sub


' ********* Center Visor Flasher **********

' FadeLights.obj(102) = array(flash102) 'center visor flash
Sub Flasherset17(value)
  if value < FlashLevel17 Then
    FlashLevel17dir =  -1
  elseif value > FlashLevel17 Then
    FlashLevel17dir = 1
  end If

  FlashLevel17 = value
  Flash102.Timerenabled = true

End Sub

Sub Flash102_timer()
  me.timerinterval = FlasherInterval

  If FlashLevel17dir > 0 Then
    If Flash2Level17 < FlashLevel17 Then
      Flash2Level17 = Flash2Level17 + FlashFadeUp
    Else
      Flash2Level17 = FlashLevel17
      me.timerenabled = False
    End If
  Else
    If Flash2Level17 > FlashLevel17 Then
      Flash2Level17 = Flash2Level17 - FlashFadeDown
    Else
      Flash2Level17 = FlashLevel17
      me.timerenabled = False
    End If
  End If

  if Flash2Level17 >  0 Then
    flash102.visible = 1
    flash102.IntensityScale = Flash2Level17 / 255
  else
    flash102.visible = 0
  end if

  if Flash2Level17 > 0 Then
    mask_prim.image="maskon"
  else
    mask_prim.image="maskoff"
  end if
End Sub

' ********* Pinbot Face Flasher **********

' FadeLights.obj(109) = array(l109, l109a, l109z, f109)
Sub Flasherset18(value)
  if value < FlashLevel18 Then
    FlashLevel18dir =  -1
  elseif value > FlashLevel18 Then
    FlashLevel18dir = 1
  end If

  FlashLevel18 = value
  F109.Timerenabled = true
End Sub

Sub F109_timer()
  me.timerinterval = FlasherInterval

  If FlashLevel18dir > 0 Then
    If Flash2Level18 < FlashLevel18 Then
      Flash2Level18 = Flash2Level18 + FlashFadeUp
    Else
      Flash2Level18 = FlashLevel18
      me.timerenabled = False
    End If
  Else
    If Flash2Level18 > FlashLevel18 Then
      Flash2Level18 = Flash2Level18 - FlashFadeDown
    Else
      Flash2Level18 = FlashLevel18
      me.timerenabled = False
    End If
  End If

  if Flash2Level18 >  0 Then
    l109.state = 1
    l109a.state = 1
    l109z.state = 1
    f109.visible = 1
    l109.IntensityScale = Flash2Level18 / 255
    l109a.IntensityScale = Flash2Level18 / 255
    l109z.IntensityScale = Flash2Level18 / 255
    f109.IntensityScale = Flash2Level18 / 255
  else
    l109.state = 0
    l109a.state = 0
    l109z.state = 0
    f109.visible = 0
  end if
End Sub

' ********* Bumper Flasher **********

' FadeLights.obj(107) = array(l107, l107a, l107z,BumpFlash,FlasherB,Flash107wall)
Sub Flasherset19(value)
  if value < FlashLevel19 Then
    FlashLevel19dir =  -1
  elseif value > FlashLevel19 Then
    FlashLevel19dir = 1
  end If

  FlashLevel19 = value
  l107.Timerenabled = true
End Sub

Sub l107_timer()
  me.timerinterval = FlasherInterval

  If FlashLevel19dir > 0 Then
    If Flash2Level19 < FlashLevel19 Then
      Flash2Level19 = Flash2Level19 + FlashFadeUp
    Else
      Flash2Level19 = FlashLevel19
      me.timerenabled = False
    End If
  Else
    If Flash2Level19 > FlashLevel19 Then
      Flash2Level19 = Flash2Level19 - FlashFadeDown
    Else
      Flash2Level19 = FlashLevel19
      me.timerenabled = False
    End If
  End If

  if Flash2Level19 >  0 Then
    l107.state = 1
    l107a.state = 1
    l107z.state = 1
    BumpFlash.visible = 1
    FlasherB.visible = 1
    Flash107wall.visible = 1
    l109.IntensityScale = Flash2Level19 / 255
    l109a.IntensityScale = Flash2Level19 / 255
    l109z.IntensityScale = Flash2Level19 / 255
    BumpFlash.IntensityScale = Flash2Level19 / 255
    FlasherB.IntensityScale = Flash2Level19 / 255
    Flash107wall.IntensityScale = Flash2Level19 / 255
  else
    l107.state = 0
    l107a.state = 0
    l107z.state = 0
    BumpFlash.visible = 0
    FlasherB.visible = 0
    Flash107wall.visible = 0
  end if
End Sub


' ********* Lower Left Flasher **********

' FadeLights.obj(106) = array(Flasher1L,Flasher1L1,FLasher1L2,Flash106wall) 'lower left flasher
Sub Flasherset20(value)
  if value < FlashLevel20 Then
    FlashLevel20dir =  -1
  elseif value > FlashLevel20 Then
    FlashLevel20dir = 1
  end If

  FlashLevel20 = value
  Flasher1L.Timerenabled = true
End Sub

Sub Flasher1L_timer()
  me.timerinterval = FlasherInterval

  If FlashLevel20dir > 0 Then
    If Flash2Level20 < FlashLevel20 Then
      Flash2Level20 = Flash2Level20 + FlashFadeUp
    Else
      Flash2Level20 = FlashLevel20
      me.timerenabled = False
    End If
  Else
    If Flash2Level20 > FlashLevel20 Then
      Flash2Level20 = Flash2Level20 - FlashFadeDown
    Else
      Flash2Level20 = FlashLevel20
      me.timerenabled = False
    End If
  End If

  if Flash2Level20 >  0 Then
    Flasher1L.visible = 1
    Flasher1L1.visible = 1
    Flasher1L2.visible = 1
    Flash106wall.visible = 1
    Flasher1L.IntensityScale = Flash2Level20 / 255
    Flasher1L1.IntensityScale = Flash2Level20 / 255
    Flasher1L2.IntensityScale = Flash2Level20 / 255
    Flash106wall.IntensityScale = Flash2Level20 / 255
  else
    Flasher1L.visible = 0
    Flasher1L1.visible = 0
    Flasher1L2.visible = 0
    Flash106wall.visible = 0
  end if

  if Flash2Level20 > 0 Then
    if Flash2Level20/255 < 0.2 then
      Flasher2.image="flash1lgion1"
      Flasher2.blenddisablelighting=0.8
    elseif Flash2Level20/255 <0.4 then
      Flasher2.image="flash1lgion2"
      Flasher2.blenddisablelighting=0.85
    elseif Flash2Level20/255 <0.6 then
      Flasher2.image="flash1lgion3"
      Flasher2.blenddisablelighting=0.9
    elseif Flash2Level20/255 <0.8 then
      Flasher2.image="flash1lgion4"
      Flasher2.blenddisablelighting=0.95
    else
      Flasher2.image="flash1lgion5"
      Flasher2.blenddisablelighting=1
    end if

    if Gi1IsOff Then
      acorns.image="acornsoffflash1"
    else
      acorns.image="acornsonflash1"
    end if
  else
    if Gi1IsOff Then
      Flasher2.image="flasher1loffgioff" 'giisoff, flasher is on (lamp 106) (same texture as when gi is on)
      Flasher2.blenddisablelighting=.8
      acorns.image="acornsoff"
    Else
      Flasher2.image="flasher1loffgion" 'giisoff, flasher is on (lamp 106) (same texture as when gi is on)
      Flasher2.blenddisablelighting=.8
      acorns.image="acornson"
    end if
  end if
End Sub


' ********* Center Left Flasher **********

' FadeLights.obj(108) = array(Flasher2L,Flasher2L1,Flasher2L2,Flash108wall)
Sub Flasherset21(value)
  if value < FlashLevel21 Then
    FlashLevel21dir =  -1
  elseif value > FlashLevel21 Then
    FlashLevel21dir = 1
  end If

  FlashLevel21 = value
  Flasher2L.Timerenabled = true
End Sub

Sub Flasher2L_timer()
  me.timerinterval = FlasherInterval

  If FlashLevel21dir > 0 Then
    If Flash2Level21 < FlashLevel21 Then
      Flash2Level21 = Flash2Level21 + FlashFadeUp
    Else
      Flash2Level21 = FlashLevel21
      me.timerenabled = False
    End If
  Else
    If Flash2Level21 > FlashLevel21 Then
      Flash2Level21 = Flash2Level21 - FlashFadeDown
    Else
      Flash2Level21 = FlashLevel21
      me.timerenabled = False
    End If
  End If

  if Flash2Level21 >  0 Then
    Flasher2L.visible = 1
    Flasher2L1.visible = 1
    Flasher2L2.visible = 1
    Flash108wall.visible = 1
    Flasher2L.IntensityScale = Flash2Level21 / 255
    Flasher2L1.IntensityScale = Flash2Level21 / 255
    Flasher2L2.IntensityScale = Flash2Level21 / 255
    Flash108wall.IntensityScale = Flash2Level21 / 255
  else
    Flasher2L.visible = 0
    Flasher2L1.visible = 0
    Flasher2L2.visible = 0
    Flash108wall.visible = 0
  end if

  if Flash2Level21 > 0 Then
    if Flash2Level21/255 < 0.2 then
      Flasher3.image="flash2gion1"
      Flasher3.blenddisablelighting=0.8
    elseif Flash2Level21/255 <0.4 then
      Flasher3.image="flash2gion2"
      Flasher3.blenddisablelighting=0.85
    elseif Flash2Level21/255 <0.6 then
      Flasher3.image="flash2gion3"
      Flasher3.blenddisablelighting=0.9
    elseif Flash2Level21/255 <0.8 then
      Flasher3.image="flash2gion4"
      Flasher3.blenddisablelighting=0.95
    else
      Flasher3.image="flash2gion5"
      Flasher3.blenddisablelighting=1
    end if
  else
    if Gi1IsOff Then
      Flasher3.image="flasher2offgioff"
      Flasher3.blenddisablelighting=.8
    Else
      Flasher3.image="flasher2offgion"
      Flasher3.blenddisablelighting=.8
    end if
  end if
End Sub

' ********* Lower Right Flasher **********

' FadeLights.obj(103) = array(Flasher1R,Flasher1R1,FLasher1R2,Flash103wall) 'lower right flasher
Sub Flasherset22(value)
  if value < FlashLevel22 Then
    FlashLevel22dir =  -1
  elseif value > FlashLevel22 Then
    FlashLevel22dir = 1
  end If

  FlashLevel22 = value
  Flasher1R.Timerenabled = true
End Sub

Sub Flasher1R_timer()
  me.timerinterval = FlasherInterval

  If FlashLevel22dir > 0 Then
    If Flash2Level22 < FlashLevel22 Then
      Flash2Level22 = Flash2Level22 + FlashFadeUp
    Else
      Flash2Level22 = FlashLevel22
      me.timerenabled = False
    End If
  Else
    If Flash2Level22 > FlashLevel22 Then
      Flash2Level22 = Flash2Level22 - FlashFadeDown
    Else
      Flash2Level22 = FlashLevel22
      me.timerenabled = False
    End If
  End If

  if Flash2Level22 >  0 Then
    Flasher1R.visible = 1
    Flasher1R1.visible = 1
    Flasher1R2.visible = 1
    Flash103wall.visible = 1
    Flasher1R.IntensityScale = Flash2Level22 / 255
    Flasher1R1.IntensityScale = Flash2Level22 / 255
    Flasher1R2.IntensityScale = Flash2Level22 / 255
    Flash103wall.IntensityScale = Flash2Level22 / 255
  else
    Flasher1R.visible = 0
    Flasher1R1.visible = 0
    Flasher1R2.visible = 0
    Flash103wall.visible = 0
  end if

  if Flash2Level22 > 0 Then
    flasher1.image="flasher1ron"
    flasher1.blenddisablelighting=1
    if Gi3IsOff Then
      metalramp_prim.image="metalrampGIOFfflash"
    Else
      metalramp_prim.image="metalrampflash"
    end if
  else
    if Gi3IsOff Then
      metalramp_prim.image="metalrampGIOFf"
    Else
      metalramp_prim.image="metalrampGIOn"
    end if
    flasher1.image="flasher1roffgion"
    flasher1.blenddisablelighting=.8
  end if
End Sub

' ********* Flasher 1 (L) **********

' FadeLights.obj(110) = array(flasherb1a,flasherb1b,Flasherplateb1,Flash110wall)
Sub Flasherset23(value)
  if value < FlashLevel23 Then
    FlashLevel23dir =  -1
  elseif value > FlashLevel23 Then
    FlashLevel23dir = 1
  end If

  FlashLevel23 = value
  Flasherb1a.Timerenabled = true
End Sub

Sub Flasherb1a_timer()
  me.timerinterval = FlasherInterval

  If FlashLevel23dir > 0 Then
    If Flash2Level23 < FlashLevel23 Then
      Flash2Level23 = Flash2Level23 + FlashFadeUp
    Else
      Flash2Level23 = FlashLevel23
      me.timerenabled = False
    End If
  Else
    If Flash2Level23 > FlashLevel23 Then
      Flash2Level23 = Flash2Level23 - FlashFadeDown
    Else
      Flash2Level23 = FlashLevel23
      me.timerenabled = False
    End If
  End If

  if Flash2Level23 >  0 Then
    flasherb1a.visible = 1
    flasherb1b.visible = 1
    Flasherplateb1.visible = 1
    Flash110wall.visible = 1
    flasherb1a.IntensityScale = Flash2Level23 / 255
    flasherb1b.IntensityScale = Flash2Level23 / 255
    Flasherplateb1.IntensityScale = Flash2Level23 / 255
    Flash110wall.IntensityScale = Flash2Level23 / 255
  else
    flasherb1a.visible = 0
    flasherb1b.visible = 0
    Flasherplateb1.visible = 0
    Flash110wall.visible = 0
  end if

  if Flash2Level23 > 0 Then
    if Flash2Level23/255 < 0.2 Then
      flasherb1.image="flashback_on1"
    elseif Flash2Level23/255 < 0.4 Then
      flasherb1.image="flashback_on2"
    elseif Flash2Level23/255 < 0.6 Then
      flasherb1.image="flashback_on3"
    elseif Flash2Level23/255 < 0.8 Then
      flasherb1.image="flashback_on4"
    else
      flasherb1.image="flashback_on5"
    end if
  else
    flasherb1.image="flashersoff"
  end if
End Sub

' ********* Flasher 2 (LC) **********

' FadeLights.obj(111) = array(flasherb2a,flasherb2b,Flasherplateb2)
Sub Flasherset24(value)
  if value < FlashLevel24 Then
    FlashLevel24dir =  -1
  elseif value > FlashLevel24 Then
    FlashLevel24dir = 1
  end If

  FlashLevel24 = value
  Flasherb2a.Timerenabled = true
End Sub

Sub Flasherb2a_timer()
  me.timerinterval = FlasherInterval

  If FlashLevel24dir > 0 Then
    If Flash2Level24 < FlashLevel24 Then
      Flash2Level24 = Flash2Level24 + FlashFadeUp
    Else
      Flash2Level24 = FlashLevel24
      me.timerenabled = False
    End If
  Else
    If Flash2Level24 > FlashLevel24 Then
      Flash2Level24 = Flash2Level24 - FlashFadeDown
    Else
      Flash2Level24 = FlashLevel24
      me.timerenabled = False
    End If
  End If

  if Flash2Level24 >  0 Then
    flasherb2a.visible = 1
    flasherb2b.visible = 1
    Flasherplateb2.visible = 1
    flasherb2a.IntensityScale = Flash2Level24 / 255
    flasherb2b.IntensityScale = Flash2Level24 / 255
    Flasherplateb2.IntensityScale = Flash2Level24 / 255
  else
    flasherb2a.visible = 0
    flasherb2b.visible = 0
    Flasherplateb2.visible = 0
  end if

  if Flash2Level24 > 0 Then
    if Flash2Level24/255 < 0.2 Then
      flasherb2.image="flashback_on1"
    elseif Flash2Level24/255 < 0.4 Then
      flasherb2.image="flashback_on2"
    elseif Flash2Level24/255 < 0.6 Then
      flasherb2.image="flashback_on3"
    elseif Flash2Level24/255 < 0.8 Then
      flasherb2.image="flashback_on4"
    else
      flasherb2.image="flashback_on5"
    end if
  else
    flasherb2.image="flashersoff"
  end if
End Sub

' ********* Flasher 3 (C) **********

' FadeLights.obj(112) = array(flasherb3a,flasherb3b,Flasherplateb3)
Sub Flasherset25(value)
  if value < FlashLevel25 Then
    FlashLevel25dir =  -1
  elseif value > FlashLevel25 Then
    FlashLevel25dir = 1
  end If

  FlashLevel25 = value
  Flasherb3a.Timerenabled = true
End Sub

Sub Flasherb3a_timer()
  me.timerinterval = FlasherInterval

  If FlashLevel25dir > 0 Then
    If Flash2Level25 < FlashLevel25 Then
      Flash2Level25 = Flash2Level25 + FlashFadeUp
    Else
      Flash2Level25 = FlashLevel25
      me.timerenabled = False
    End If
  Else
    If Flash2Level25 > FlashLevel25 Then
      Flash2Level25 = Flash2Level25 - FlashFadeDown
    Else
      Flash2Level25 = FlashLevel25
      me.timerenabled = False
    End If
  End If

  if Flash2Level25 >  0 Then
    flasherb3a.visible = 1
    flasherb3b.visible = 1
    Flasherplateb3.visible = 1
    flasherb3a.IntensityScale = Flash2Level25 / 255
    flasherb3b.IntensityScale = Flash2Level25 / 255
    Flasherplateb3.IntensityScale = Flash2Level25 / 255
  else
    flasherb3a.visible = 0
    flasherb3b.visible = 0
    Flasherplateb3.visible = 0
  end if

  if Flash2Level25 > 0 Then
    if Flash2Level25/255 < 0.2 Then
      flasherb3.image="flashback_on1"
    elseif Flash2Level25/255 < 0.4 Then
      flasherb3.image="flashback_on2"
    elseif Flash2Level25/255 < 0.6 Then
      flasherb3.image="flashback_on3"
    elseif Flash2Level25/255 < 0.8 Then
      flasherb3.image="flashback_on4"
    else
      flasherb3.image="flashback_on5"
    end if
  else
    flasherb3.image="flashersoff"
  end if
End Sub


' ********* Flasher 4 (RC) **********

' FadeLights.obj(113) = array(flasherb4a,flasherb4b,Flasherplateb4)
Sub Flasherset26(value)
  if value < FlashLevel26 Then
    FlashLevel26dir =  -1
  elseif value > FlashLevel26 Then
    FlashLevel26dir = 1
  end If

  FlashLevel26 = value
  Flasherb4a.Timerenabled = true
End Sub

Sub Flasherb4a_timer()
  me.timerinterval = FlasherInterval

  If FlashLevel26dir > 0 Then
    If Flash2Level26 < FlashLevel26 Then
      Flash2Level26 = Flash2Level26 + FlashFadeUp
    Else
      Flash2Level26 = FlashLevel26
      me.timerenabled = False
    End If
  Else
    If Flash2Level26 > FlashLevel26 Then
      Flash2Level26 = Flash2Level26 - FlashFadeDown
    Else
      Flash2Level26 = FlashLevel26
      me.timerenabled = False
    End If
  End If

  if Flash2Level26 >  0 Then
    flasherb4a.visible = 1
    flasherb4b.visible = 1
    Flasherplateb4.visible = 1
    flasherb4a.IntensityScale = Flash2Level26 / 255
    flasherb4b.IntensityScale = Flash2Level26 / 255
    Flasherplateb4.IntensityScale = Flash2Level26 / 255
  else
    flasherb4a.visible = 0
    flasherb4b.visible = 0
    Flasherplateb4.visible = 0
  end if

  if Flash2Level26 > 0 Then
    if Flash2Level26/255 < 0.2 Then
      flasherb4.image="flashback_on1"
    elseif Flash2Level26/255 < 0.4 Then
      flasherb4.image="flashback_on2"
    elseif Flash2Level26/255 < 0.6 Then
      flasherb4.image="flashback_on3"
    elseif Flash2Level26/255 < 0.8 Then
      flasherb4.image="flashback_on4"
    else
      flasherb4.image="flashback_on5"
    end if
  else
    flasherb4.image="flashersoff"
  end if
End Sub


' ********* Flasher 5 (R) **********

' FadeLights.obj(114) = array(flasherb5a,flasherb5b,Flasherplateb5,Flash114wall)
Sub Flasherset27(value)
  if value < FlashLevel27 Then
    FlashLevel27dir =  -1
  elseif value > FlashLevel27 Then
    FlashLevel27dir = 1
  end If

  FlashLevel27 = value
  Flasherb5a.Timerenabled = true
End Sub

Sub Flasherb5a_timer()
  me.timerinterval = FlasherInterval

  If FlashLevel27dir > 0 Then
    If Flash2Level27 < FlashLevel27 Then
      Flash2Level27 = Flash2Level27 + FlashFadeUp
    Else
      Flash2Level27 = FlashLevel27
      me.timerenabled = False
    End If
  Else
    If Flash2Level27 > FlashLevel27 Then
      Flash2Level27 = Flash2Level27 - FlashFadeDown
    Else
      Flash2Level27 = FlashLevel27
      me.timerenabled = False
    End If
  End If

  dim flashx3 : flashx3 = (FlashLevel27/255)

  if flashx3 >  0 Then
    flasherb5a.visible = 1
    flasherb5b.visible = 1
    Flasherplateb5.visible = 1
    Flash114wall.visible = 1
    flasherb5a.IntensityScale = flashx3
    flasherb5b.IntensityScale = flashx3
    Flasherplateb5.IntensityScale = flashx3
    Flash114wall.IntensityScale = flashx3
  else
    flasherb5a.visible = 0
    flasherb5b.visible = 0
    Flasherplateb5.visible = 0
    Flash114wall.visible = 0
  end if

  if flashx3 > 0 Then
    if flashx3 < 0.2 Then
      flasherb5.image="flashback_on1"
    elseif flashx3 < 0.4 Then
      flasherb5.image="flashback_on2"
    elseif flashx3 < 0.6 Then
      flasherb5.image="flashback_on3"
    elseif flashx3 < 0.8 Then
      flasherb5.image="flashback_on4"
    else
      flasherb5.image="flashback_on5"
    end if
  else
    flasherb5.image="flashersoff"
  end if
End Sub


'*****************************************
'         General Illumination
'*****************************************
Set GiCallBack2 = GetRef("UpdateGi")

Sub UpdateGi(nr,step)
  Dim ii
  Select Case nr

  Case 0    'Bottom Playfield
    If step=0 Then
      setlamp 115,0
      setlamp 116,1
      For each ii in GI_Bottom:ii.state=0:Next
      gi0isoff=true
      For each ii in rubber_bottom: ii.image = "rubbasGIOFF": Next
      laneguideL.image="metallaneguideloff"
      laneguideR.image="metallaneguideroff"
      acorns.image="acornsoff"
      metals.image="metalpostsoff"
      outerinner_prim.image="woodguideoff"
      apronrail.image="apronrailoff"
      rightsign.image="rightsignoff"
    Else
      setlamp 115,1
      setlamp 116,0
      For each ii in GI_Bottom:ii.state=1:Next
      gi0isoff=false
      For each ii in rubber_bottom: ii.image = "rubbasGION": Next
      laneguideL.image="metallaneguidelon"
      laneguideR.image="metallaneguideron"
      acorns.image="acornson"
      metals.image="metalpostson"
      outerinner_prim.image="woodguideon"
      apronrail.image="apronrailon"
      rightsign.image="rightsignon"
    End If
    For each ii in GI_Bottom:ii.IntensityScale = 0.125 * step:Next

  Case 1    'Left Playfield
    If step=0 Then
      For each ii in GI_Left:ii.state=0:Next
      outerl.image="outerleftoff"
      gi1isoff=true
      sw18.image="drop1off"
      sw17.image="drop1off"
      sw16.image="drop1off"
      outerl.image="outerleftoff"
      Flasher1L.timerenabled=1'flasher2.image="flasher1loffgioff"
      Flasher2L.timerenabled=1'flasher3.image="flasher2offgioff"
      For each ii in rubber_left: ii.image = "rubbasGIOFF": Next
    Else
      For each ii in GI_Left:ii.state=1:Next
      outerl.image="outerlefton"
      gi1isoff=false
      sw18.image="drop3on"
      sw17.image="drop2on"
      sw16.image="drop1on"
      outerl.image="outerlefton"
      Flasher1L.timerenabled=1'flasher2.image="flasher1loffgion"
      Flasher2L.timerenabled=1'flasher3.image="flasher2offgion"
      For each ii in rubber_left: ii.image = "rubbasGION": Next
    End If
    For each ii in GI_Left:ii.IntensityScale = 0.125 * step:Next

  Case 2    'Upper Playfield
    If step=0 Then
      For each ii in GI_Upper:ii.state=0:Next
      gi2isoff=true
      For each ii in rubber_upper: ii.image = "rubbasGIOFF": Next
      rubberposts.image="rubberpostsoff"
      helmet_prim.blenddisablelighting=0.05
      targetwrapper_prim.blenddisablelighting=0.05
      corkscrew_prim.image="corkscrewoff"
      Prim_upperPF.image="upperPFGIOFF"
    Else
      For each ii in GI_Upper:ii.state=1:Next
      gi2isoff=false
      For each ii in rubber_upper: ii.image = "rubbasGION": Next
      rubberposts.image="rubberpostson"
      helmet_prim.blenddisablelighting=.25
      targetwrapper_prim.blenddisablelighting=0.14
      corkscrew_prim.image="corkscrewon"
      Prim_upperPF.image="upperPFGION"
    End If
    For each ii in GI_Upper:ii.IntensityScale = 0.125 * step:Next

  Case 3    'Right Playfield
    If step=0 Then
      For each ii in GI_Right:ii.state=0:Next
      gi3isoff=true
      sw41_prim.blenddisablelighting=0
      sw42_prim.blenddisablelighting=0
      sw43_prim.blenddisablelighting=0
      sw44_prim.blenddisablelighting=0
      sw45_prim.blenddisablelighting=0
      flasher1.image="flasher1roffgioff"
      For each ii in rubber_right: ii.image = "rubbasGIOFF": Next
      upperrubbers.image="upperpfrubbersoff"
      metalramp_prim.image="metalrampGIOFF"
      outerr.image="outerrightoff"
      r18_prim.image="18off"
    Else
      For each ii in GI_Right:ii.state=1:Next
      gi3isoff=false
      sw41_prim.blenddisablelighting=0.1
      sw42_prim.blenddisablelighting=0.1
      sw43_prim.blenddisablelighting=0.1
      sw44_prim.blenddisablelighting=0.1
      sw45_prim.blenddisablelighting=0.1
      flasher1.image="flasher1roffgion"
      For each ii in rubber_right: ii.image = "rubbasGION": Next
      upperrubbers.image="upperpfrubberson"
      metalramp_prim.image="metalrampGION"
      outerr.image="outerrighton"
      r18_prim.image="18on"
    End If
    For each ii in GI_Right:ii.IntensityScale = 0.125 * step:Next
  End Select
  UpdateLamp71
  UpdateLamp72
  UpdateLamp73
  UpdateLamp74
  UpdateLamp75
  UpdateLamp76

' Start WJR Comment
' UpdateLamp101
' UpdateLamp102
' UpdateLamp103
' UpdateLamp104
''  UpdateLamp105
' UpdateLamp106
''  UpdateLamp107
' UpdateLamp108
''  UpdateLamp109
' UpdateLamp110
' UpdateLamp111
' UpdateLamp112
' UpdateLamp113
' UpdateLamp114
' End WJR Comment

End Sub

'**********************************************************************************************************

'Solenoid Controlled toys
'**********************************************************************************************************


'Sub RampCallback(lvl)
' lramp_prim.rotx=-12 * lvl
' lrampmetal_prim.rotx=-12 * lvl
' lrampflap_prim.rotx=-12 * lvl
' lramplever_prim.transy=-40 * lvl
' lramplever_prim.transz=-16 * lvl
' If lvl < 0.33 Then
'   rightramp.Collidable = False
'   rightrampmid.Collidable = False
'   rightrampup.Collidable = true
'   lramp_prim1.visible = False
'   lramp_prim2.visible = True
' ElseIf lvl < 0.8 Then
'   rightramp.Collidable = False
'   rightrampmid.Collidable = True
'   rightrampup.Collidable = True
' Else
'   rightramp.Collidable = True
'   rightrampmid.Collidable = False
'   rightrampup.Collidable = False
'   lramp_prim1.visible = True
'   lramp_prim2.visible = False
' End If
'End Sub

Dim RampDir, RampLvl, RampSpeed
RampLvl = 1
RampSpeed = 0.05

Sub RampTimer_Timer()
  RampLvl = RampLvl + RampDir

  If RampLvl < 0 and RampDir < 0 Then
    RampLvl = 0
    me.enabled = False
  ElseIf RampLvl > 1 and RampDir > 0 Then
    RampLvl = 1
    me.enabled = False
  ElseIf RampLvl < 0.33 Then
    rightramp.Collidable = False
    rightrampmid.Collidable = False
    rightrampup.Collidable = true
    lramp_prim1.visible = True
    lramp_prim2.visible = False
    lramp_prim3.visible = False
    Controller.Switch(15)=0
  ElseIf RampLvl < 0.8 Then
    rightramp.Collidable = False
    rightrampmid.Collidable = True
    rightrampup.Collidable = True
    lramp_prim1.visible = False
    lramp_prim2.visible = True
    lramp_prim3.visible = False
  Else
    rightramp.Collidable = True
    rightrampmid.Collidable = False
    rightrampup.Collidable = False
    lramp_prim1.visible = False
    lramp_prim2.visible = False
    lramp_prim3.visible = True
    Controller.Switch(15)=1
  End If


  lramp_prim.rotx=-12 * RampLvl
  lrampmetal_prim.rotx=-12 * RampLvl
  lrampflap_prim.rotx=-12 * RampLvl
  lramplever_prim.transy=-40 * RampLvl
  lramplever_prim.transz=-16 * RampLvl
End Sub


Sub solRampUp(Enabled)
  If Enabled then
    Playsoundat "topdiverteron", lramp_prim
    'Controller.Switch(44)=0
    'SetLamp 120, 0
    RampDir = -RampSpeed
    RampTimer.enabled = True


    Dim BOT, b
    BOT = GetBalls

    If UBound(BOT) = -1 Then Exit Sub

    For b = 0 to UBound(BOT)
      If lrampflap_prim.rotx < -10 and InRect(BOT(b).x, BOT(b).y, 74, 640, 165,610,213,746,118,778) Then
        BOT(b).velz = 40
        BOT(b).velx = BOT(b).velx + 1.5
        BOT(b).vely = BOT(b).vely + 5
      Elseif lrampflap_prim.rotx < -10 and InRect(BOT(b).x, BOT(b).y, 25, 505, 117,475,165,610,74,640) Then
        BOT(b).velz = 30
        BOT(b).velx = BOT(b).velx +1
        BOT(b).vely = BOT(b).vely + 3
      Elseif lrampflap_prim.rotx < -10 and InRect(BOT(b).x, BOT(b).y, 35, 503, 30,451,118,446,120,477) Then
        BOT(b).velz = 15
        BOT(b).velx = BOT(b).velx + 0.5
        BOT(b).vely = BOT(b).vely + 1.5
      End If
    Next
  end if
End Sub

Sub solRampDwn(Enabled)
  If Enabled then
    'Controller.Switch(44)=1
    'SetLamp 120, 1
    RampDir = RampSpeed
    RampTimer.enabled = True
    Playsoundat "topdiverteroff", lramp_prim
  end if
End Sub

 'VISOR
Sub UpdateVisor(currpos, currspeed, lastpos)
  dim xx
  For each xx in Visor:xx.RotX  = currpos * 12 / 58 : Next
  For each xx in targetbank:xx.TransZ  = -currpos: Next
  if currpos > 50 Then
    targetslow.collidable = false
  elseif currpos > 42 then
    targetslow.collidable = true
    targetsmid.collidable = false
  elseif currpos > 14 then
    targetsmid.collidable = true
    targetslow.collidable = false
    targetshigh.collidable=false
    for each xx in targetbanktargets:xx.Collidable = false:Next
  elseif currpos < 14 Then
    targetsmid.collidable=false
    targetshigh.collidable=true
    for each xx in targetbanktargets:xx.Collidable = true:Next
  end if
  xx = CInt((currpos * 9) / 29)
  if xx > 8 then xx = 8
  targetwrapper_prim.image = "targetbank" & chr(97+xx)
  'debug.print currpos
  If currspeed <> 0 then
    PlaySound SoundFX("gunmotor",DOFGear), 0, 1, AudioPan(sw43), 0,0, 1, 0, AudioFade( sw43)
  Else
    Stopsound "gunmotor"
  end if
 End Sub


'**********************************************************************************************************

 ' Drain hole and kickers
Sub Drain_Hit:bsTrough.addball me : playsoundat"drain", drain : End Sub
Sub sw47_Hit:bsLEye.AddBall 0 : playsoundat "kicker_enter_center", sw47: End Sub
Sub sw48_Hit:bsREye.AddBall 0 : playsoundat "kicker_enter_center", sw48: End Sub
Sub sw46_Hit:bsSaucer.AddBall 0 : playsoundat "kicker_enter_center", sw46: End Sub

dim TargetTransY, TargetDelay
TargetTransY = -8
TargetDelay = 75

'Standup targets
Sub sw67_Hit:vpmTimer.PulseSw 67:playsoundat"target", sw67:End Sub
Sub sw41_Hit:vpmTimer.PulseSw 41:playsoundat"target", sw41:sw41_prim.transy=TargetTransY:vpmTimer.AddTimer TargetDelay,"sw41_prim.transy=TargetTransY/2'":vpmTimer.AddTimer TargetDelay+TargetDelay,"sw41_prim.transy=0'":End Sub
Sub sw42_Hit:vpmTimer.PulseSw 42:playsoundat"target", sw42:sw42_prim.transy=TargetTransY:vpmTimer.AddTimer TargetDelay,"sw42_prim.transy=TargetTransY/2'":vpmTimer.AddTimer TargetDelay+TargetDelay,"sw42_prim.transy=0'":End Sub
Sub sw43_Hit:vpmTimer.PulseSw 43:playsoundat"target", sw43:sw43_prim.transy=TargetTransY:vpmTimer.AddTimer TargetDelay,"sw43_prim.transy=TargetTransY/2'":vpmTimer.AddTimer TargetDelay+TargetDelay,"sw43_prim.transy=0'":End Sub
Sub sw44_Hit:vpmTimer.PulseSw 44:playsoundat"target", sw44:sw44_prim.transy=TargetTransY:vpmTimer.AddTimer TargetDelay,"sw44_prim.transy=TargetTransY/2'":vpmTimer.AddTimer TargetDelay+TargetDelay,"sw44_prim.transy=0'":End Sub
Sub sw45_Hit:vpmTimer.PulseSw 45:playsoundat"target", sw45:sw45_prim.transy=TargetTransY:vpmTimer.AddTimer TargetDelay,"sw45_prim.transy=TargetTransY/2'":vpmTimer.AddTimer TargetDelay+TargetDelay,"sw45_prim.transy=0'":End Sub
Sub sw51_Hit:vpmTimer.PulseSw 51:playsoundat"target", sw51:sw51_prim.transy=TargetTransY:vpmTimer.AddTimer TargetDelay,"sw51_prim.transy=TargetTransY/2'":vpmTimer.AddTimer TargetDelay+TargetDelay,"sw51_prim.transy=0'":End Sub
Sub sw52_Hit:vpmTimer.PulseSw 52:playsoundat"target", sw52:sw52_prim.transy=TargetTransY:vpmTimer.AddTimer TargetDelay,"sw52_prim.transy=TargetTransY/2'":vpmTimer.AddTimer TargetDelay+TargetDelay,"sw52_prim.transy=0'":End Sub
Sub sw53_Hit:vpmTimer.PulseSw 53:playsoundat"target", sw53:sw53_prim.transy=TargetTransY:vpmTimer.AddTimer TargetDelay,"sw53_prim.transy=TargetTransY/2'":vpmTimer.AddTimer TargetDelay+TargetDelay,"sw53_prim.transy=0'":End Sub
Sub sw54_Hit:vpmTimer.PulseSw 54:playsoundat"target", sw54:sw54_prim.transy=TargetTransY:vpmTimer.AddTimer TargetDelay,"sw54_prim.transy=TargetTransY/2'":vpmTimer.AddTimer TargetDelay+TargetDelay,"sw54_prim.transy=0'":End Sub
Sub sw55_Hit:vpmTimer.PulseSw 55:playsoundat"target", sw55:sw55_prim.transy=TargetTransY:vpmTimer.AddTimer TargetDelay,"sw55_prim.transy=TargetTransY/2'":vpmTimer.AddTimer TargetDelay+TargetDelay,"sw55_prim.transy=0'":End Sub
Sub sw38_Hit:vpmTimer.PulseSw 38:playsoundat"target", sw38:End Sub

'Combo Target Hits
Sub sw4142_Hit:sw41_Hit:sw42_Hit:End Sub
Sub sw4243_Hit:sw42_Hit:sw43_Hit:End Sub
Sub sw4344_Hit:sw43_Hit:sw44_Hit:End Sub
Sub sw4445_Hit:sw44_Hit:sw45_Hit:End Sub

Sub sw5152_Hit:sw51_Hit:sw52_Hit:End Sub
Sub sw5253_Hit:sw52_Hit:sw53_Hit:End Sub
Sub sw5354_Hit:sw53_Hit:sw54_Hit:End Sub
Sub sw5455_Hit:sw54_Hit:sw55_Hit:End Sub

'Drop Targets
Sub Sw18_Dropped:dtDTBank.Hit 1:sw1718.collidable=false:End Sub
Sub sw17_Dropped:dtDTBank.Hit 2:sw1718.collidable=false:sw1617.collidable=false:End Sub
Sub sw16_Dropped:dtDTBank.Hit 3:sw1617.collidable=false:End Sub

Sub Sw18_Raised:sw1718.collidable=true:sw1617.collidable=true:End Sub
Sub Sw17_Raised:sw1718.collidable=true:sw1617.collidable=true:End Sub
Sub Sw16_Raised:sw1718.collidable=true:sw1617.collidable=true:End Sub

'Combo Drop Target
Sub sw1617_HIt:sw16_Dropped:sw17_dropped:End Sub
Sub sw1718_HIt:sw18_Dropped:sw17_dropped:End Sub

'Wire Triggers
Sub sw25_Hit:Controller.Switch(25) = 1:PlaySoundat"rollover", sw25:End Sub
Sub sw25_UnHit:Controller.Switch(25) = 0:End Sub
Sub sw26_Hit:Controller.Switch(26) = 1:PlaySoundat"rollover", sw26:End Sub
Sub sw26_UnHit:Controller.Switch(26) = 0:End Sub
Sub sw27_Hit:Controller.Switch(27) = 1:PlaySoundat"rollover", sw27:End Sub
Sub sw27_UnHit:Controller.Switch(27) = 0:End Sub
Sub sw28_Hit:Controller.Switch(28) = 1:PlaySoundat"rollover", sw28:End Sub
Sub sw28_UnHit:Controller.Switch(28) = 0:End Sub
Sub sw56_Hit:me.TimerInterval=100:me.TimerEnabled = 1:PlaySoundat "rollover", sw56:End Sub  'Debug.Print "sw56":
Sub sw56_Timer:vpmTimer.PulseSw 56:me.TimerEnabled = 0:End Sub
Sub sw57_Hit:vpmTimer.PulseSw 57:PlaySoundat"rollover", sw57:End Sub  'Debug.Print "sw57":
Sub sw58_Hit:Controller.Switch(58) = 1:PlaySoundat"rollover", sw58:End Sub
Sub sw58_UnHit:Controller.Switch(58) = 0:End Sub
Sub rampout_Hit:Controller.Switch(36) = 1:End Sub
Sub rampout_UnHit:Controller.Switch(36) = 0:End Sub
Sub rampin_Hit:Controller.Switch(37) = 1:End Sub
Sub rampin_UnHit:Controller.Switch(37) = 0:End Sub
Sub sw68_Hit:Controller.Switch(68) = 1:End Sub
Sub sw68_UnHit:Controller.Switch(68) = 0:End Sub


'Bumpers
Sub sw62_bumper2_Hit : vpmTimer.PulseSw(62) : playsoundat SoundFX("fx_bumper1",DOFContactors), sw62_bumper2: End Sub
Sub sw61_bumper1_Hit : vpmTimer.PulseSw(61) : playsoundat SoundFX("fx_bumper2",DOFContactors), sw61_bumper1: End Sub
Sub sw63_bumper3_Hit : vpmTimer.PulseSw(63) : playsoundat SoundFX("fx_bumper3",DOFContactors), sw63_bumper3: End Sub

'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep, LaStep, LaaStep, LbStep, LcStep, LccStep, LcccStep, RaStep, RbStep, RcStep, RccStep, RcccStep, ReStep, RfStep, RffStep, RgStep, RhStep, RiStep, RjStep, RkStep

Sub RightSlingShot_Slingshot
  vpmTimer.PulseSw 64
    PlaySound SoundFX("right_slingshot",DOFContactors), 0,1, 0.05,0.05 '0,1, AudioPan(RightSlingShot), 0.05,0,0,1,AudioFade(RightSlingShot)
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.TransZ = -29
    RStep = 0
    RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.TransZ = -12
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.TransZ = 0:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
  vpmTimer.PulseSw 65
    PlaySound SoundFX("left_slingshot",DOFContactors), 0,1, -0.05,0.05 '0,1, AudioPan(LeftSlingShot), 0.05,0,0,1,AudioFade(LeftSlingShot)
    LSling.Visible = 0
    LSling1.Visible = 1
    sling2.TransZ = -31
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.TransZ = -16
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.TransZ = 0:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

Sub WallLA_Hit
    RubberLA.Visible = 0
    RubberLA3.Visible = 1
    LAstep = 0
    WallLA.TimerEnabled = 1
End Sub

Sub WallLA_Timer
    Select Case LAstep
        Case 3:RubberLA3.Visible = 0:RubberLA4.Visible = 1
        Case 4:RubberLA4.Visible = 0:RubberLA.Visible = 1:WallLA.TimerEnabled = 0:
    End Select
    LAstep = LAstep + 1
End Sub

Sub WallLAa_Hit
  vpmTimer.PulseSw 11
    RubberLA.Visible = 0
    RubberLA1.Visible = 1
    LAastep = 0
    WallLAa.TimerEnabled = 1
End Sub

Sub WallLAa_Timer
    Select Case LAastep
        Case 3:RubberLA1.Visible = 0:RubberLA2.Visible = 1
        Case 4:RubberLA2.Visible = 0:RubberLA.Visible = 1:WallLAa.TimerEnabled = 0:
    End Select
    LAastep = LAastep + 1
End Sub

Sub WallLb_Hit
    RubberLb.Visible = 0
    RubberLb1.Visible = 1
    Lbstep = 0
    WallLb.TimerEnabled = 1
End Sub

Sub WallLb_Timer
    Select Case Lbstep
        Case 3:RubberLb1.Visible = 0:RubberLb2.Visible = 1
        Case 4:RubberLb2.Visible = 0:RubberLb.Visible = 1:WallLb.TimerEnabled = 0:
    End Select
    Lbstep = Lbstep + 1
End Sub

Sub WallLc_Hit
    RubberLc.Visible = 0
    RubberLc1.Visible = 1
    Lcstep = 0
    WallLc.TimerEnabled = 1
End Sub

Sub WallLc_Timer
    Select Case Lcstep
        Case 3:RubberLc1.Visible = 0:RubberLc2.Visible = 1
        Case 4:RubberLc2.Visible = 0:RubberLc.Visible = 1:WallLc.TimerEnabled = 0:
    End Select
    Lcstep = Lcstep + 1
End Sub

Sub Walllcc_Hit
  vpmTimer.PulseSw 12
    RubberLc.Visible = 0
    RubberLc3.Visible = 1
    lccstep = 0
    Walllcc.TimerEnabled = 1
End Sub

Sub Walllcc_Timer
    Select Case lccstep
        Case 3:RubberLc3.Visible = 0:RubberLc4.Visible = 1
        Case 4:RubberLc4.Visible = 0:RubberLc.Visible = 1:WallLc.TimerEnabled = 0:
    End Select
    lccstep = lccstep + 1
End Sub

Sub Walllccc_Hit
  vpmTimer.PulseSw 12
    RubberLc.Visible = 0
    RubberLc5.Visible = 1
    lcccstep = 0
    Walllccc.TimerEnabled = 1
End Sub

Sub Walllccc_Timer
    Select Case lcccstep
        Case 3:RubberLc5.Visible = 0:RubberLc6.Visible = 1
        Case 4:RubberLc6.Visible = 0:RubberLc.Visible = 1:WallLc.TimerEnabled = 0:
    End Select
    lcccstep = lcccstep + 1
End Sub

Sub Wallra_Hit
    Rubberra.Visible = 0
    Rubberra1.Visible = 1
    rastep = 0
    Wallra.TimerEnabled = 1
End Sub

Sub Wallra_Timer
    Select Case rastep
        Case 3:Rubberra1.Visible = 0:Rubberra2.Visible = 1
        Case 3:Rubberra2.Visible = 0:Rubberra3.Visible = 1
        Case 4:Rubberra3.Visible = 0:Rubberra.Visible = 1:Wallra.TimerEnabled = 0:
    End Select
    rastep = rastep + 1
End Sub

Sub Wallrb_Hit
  vpmTimer.PulseSw 66
    Rubberrb.Visible = 0
    Rubberrb1.Visible = 1
    rbstep = 0
    Wallrb.TimerEnabled = 1
End Sub

Sub Wallrb_Timer
    Select Case rbstep
        Case 3:Rubberrb1.Visible = 0:Rubberrb2.Visible = 1
        Case 3:Rubberrb2.Visible = 0:Rubberrb3.Visible = 1
        Case 4:Rubberrb3.Visible = 0:Rubberrb.Visible = 1:Wallrb.TimerEnabled = 0:
    End Select
    rbstep = rbstep + 1
End Sub

Sub Wallrc_Hit
    Rubberrc.Visible = 0
    Rubberrc1.Visible = 1
    rcstep = 0
    Wallrc.TimerEnabled = 1
End Sub

Sub Wallrc_Timer
    Select Case rcstep
        Case 3:Rubberrc1.Visible = 0:Rubberrc2.Visible = 1
        Case 4:Rubberrc2.Visible = 0:Rubberrc.Visible = 1:Wallrc.TimerEnabled = 0:
    End Select
    rcstep = rcstep + 1
End Sub

Sub Wallrcc_Hit
    Rubberrc.Visible = 0
    Rubberrc3.Visible = 1
    rccstep = 0
    Wallrcc.TimerEnabled = 1
End Sub

Sub Wallrcc_Timer
    Select Case rccstep
        Case 3:Rubberrc3.Visible = 0:Rubberrc4.Visible = 1
        Case 4:Rubberrc4.Visible = 0:Rubberrc.Visible = 1:Wallrc.TimerEnabled = 0:
    End Select
    rccstep = rccstep + 1
End Sub

Sub Wallrccc_Hit
    Rubberrc.Visible = 0
    Rubberrc5.Visible = 1
    rcccstep = 0
    Wallrccc.TimerEnabled = 1
End Sub

Sub Wallrccc_Timer
    Select Case rcccstep
        Case 3:Rubberrc5.Visible = 0:Rubberrc6.Visible = 1
        Case 4:Rubberrc6.Visible = 0:Rubberrc.Visible = 1:Wallrc.TimerEnabled = 0:
    End Select
    rcccstep = rcccstep + 1
End Sub

Sub Wallre_Hit
    Rubberre.Visible = 0
    Rubberre1.Visible = 1
    restep = 0
    Wallre.TimerEnabled = 1
End Sub

Sub Wallre_Timer
    Select Case restep
        Case 3:Rubberre1.Visible = 0:Rubberre2.Visible = 1
        Case 3:Rubberre2.Visible = 0:Rubberre3.Visible = 1
        Case 4:Rubberre3.Visible = 0:Rubberre.Visible = 1:Wallre.TimerEnabled = 0:
    End Select
    restep = restep + 1
End Sub

Sub Wallrf_Hit
    Rubberrf.Visible = 0
    Rubberrf1.Visible = 1
    rfstep = 0
    Wallrf.Timerenabled = 1
End Sub

Sub Wallrf_Timer
    Select Case rfstep
        Case 3:Rubberrf1.Visible = 0:Rubberrf2.Visible = 1
        Case 3:Rubberrf2.Visible = 0:Rubberrf3.Visible = 1
        Case 4:Rubberrf3.Visible = 0:Rubberrf.Visible = 1:Wallrf.Timerenabled = 0:
    End Select
    rfstep = rfstep + 1
End Sub

Sub Wallrg_Hit
    Rubberrg.Visible = 0
    Rubberrg1.Visible = 1
    rgstep = 0
    Wallrg.TimerEnabled = 1
End Sub

Sub Wallrg_Timer
    Select Case rgstep
        Case 3:Rubberrg1.Visible = 0:Rubberrg2.Visible = 1
        Case 4:Rubberrg2.Visible = 0:Rubberrg.Visible = 1:Wallrg.TimerEnabled = 0:
    End Select
    rgstep = rgstep + 1
End Sub

Sub Wallrh_Hit
    Rubberrh.Visible = 0
    Rubberrh1.Visible = 1
    rhstep = 0
    Wallrh.TimerEnabled = 1
End Sub

Sub Wallrh_Timer
    Select Case rhstep
        Case 3:Rubberrh1.Visible = 0:Rubberrh2.Visible = 1
        Case 4:Rubberrh2.Visible = 0:Rubberrh.Visible = 1:Wallrh.TimerEnabled = 0:
    End Select
    rhstep = rhstep + 1
End Sub

Sub Wallri_Hit
    Rubberri.Visible = 0
    Rubberri1.Visible = 1
    ristep = 0
    Wallri.TimerEnabled = 1
End Sub

Sub Wallri_Timer
    Select Case ristep
        Case 3:Rubberri1.Visible = 0:Rubberri2.Visible = 1
        Case 4:Rubberri2.Visible = 0:Rubberri.Visible = 1:Wallri.TimerEnabled = 0:
    End Select
    ristep = ristep + 1
End Sub

Sub Wallrj_Hit
    Rubberrj.Visible = 0
    Rubberrj1.Visible = 1
    rjstep = 0
    Wallrj.TimerEnabled = 1
End Sub

Sub Wallrj_Timer
    Select Case rjstep
        Case 3:Rubberrj1.Visible = 0:Rubberrj2.Visible = 1
        Case 4:Rubberrj2.Visible = 0:Rubberrj.Visible = 1:Wallrj.TimerEnabled = 0:
    End Select
    rjstep = rjstep + 1
End Sub

Sub Wallrk_Hit
    Rubberrk.Visible = 0
    Rubberrk1.Visible = 1
    rkstep = 0
    Wallrk.TimerEnabled = 1
End Sub

Sub Wallrk_Timer
    Select Case rkstep
        Case 3:Rubberrk1.Visible = 0:Rubberrk2.Visible = 1
        Case 4:Rubberrk2.Visible = 0:Rubberrk.Visible = 1:Wallrk.TimerEnabled = 0:
    End Select
    rkstep = rkstep + 1
End Sub

'***************************************************

InitLamps()             ' turn off the lights and flashers and reset them to the default parameters
FrameTimer.Interval = -1 'lamp fading speed
FrameTimer.Enabled = 1

' Called once per frame (16.6ms at 60hz)
Sub FrameTimer_Timer()
    Dim chgLamp, num, chg, ii
    chgLamp = Controller.ChangedLamps
    If Not IsEmpty(chgLamp) Then
        For ii = 0 To UBound(chgLamp)
       SetLamp ChgLamp(ii,0), ChgLamp(ii,1)
        Next
    End If
  FadeLights.Update
  FlipperUpdate
  BallShadowUpdate
End Sub

' Called every 1ms.
Sub OneMsec_Timer()
  FadeLights.Update1
End Sub

' div lamp subs

Sub InitLamps()
  UpdateGI 0,0:UpdateGI 1,0:UpdateGI 2,0:UpdateGI 3,0
  'FadeLights.Callback(120)="RampCallback "
  FadeLights.obj(11) = array(l11, l11a)
  FadeLights.obj(12) = array(l12, l12a)
  FadeLights.obj(13) = array(l13, l13a)
  FadeLights.obj(14) = array(l14, l14a)
  FadeLights.obj(15) = array(l15, l15a)
  FadeLights.obj(16) = array(l16, l16a)
  FadeLights.obj(17) = array(l17, l17a,f17)
  FadeLights.obj(18) = array(l18, l18a)
  FadeLights.obj(21) = array(l21, l21a)
  FadeLights.obj(22) = array(l22, l22a)
  FadeLights.obj(23) = array(l23, l23a)
  FadeLights.obj(24) = array(l24, l24a)
  FadeLights.obj(25) = array(l25, l25a)
  FadeLights.obj(26) = array(l26, l26a)
  FadeLights.obj(27) = array(l27, l27a)
  FadeLights.obj(28) = array(l28, l28a)
  FadeLights.obj(31) = array(l31, l31a)
  FadeLights.obj(32) = array(l32, l32a)
  FadeLights.obj(33) = array(l33, l33a)
  FadeLights.obj(34) = array(l34, l34a)
  FadeLights.obj(35) = array(l35, l35a)
  FadeLights.obj(36) = array(l36, l36a)
  FadeLights.obj(37) = array(l37, l37a)
  FadeLights.obj(38) = array(l38, l38a)
  FadeLights.obj(41) = array(l41, l41a)
  FadeLights.obj(42) = array(l42, l42a)
  FadeLights.obj(43) = array(l43, l43a)
  FadeLights.obj(44) = array(l44, l44a)
  FadeLights.obj(45) = array(l45, l45a)
  FadeLights.obj(46) = array(l46, l46a)
  FadeLights.obj(47) = array(l47, l47a)
  FadeLights.obj(48) = array(l48, l48a)
  FadeLights.obj(51) = array(l51, l51a)
  FadeLights.obj(52) = array(l52, l52a)
  FadeLights.obj(53) = array(l53, l53a)
  FadeLights.obj(54) = array(l54, l54a)
  FadeLights.obj(55) = array(l55, l55a)
  FadeLights.obj(56) = array(l56, l56a)
  FadeLights.obj(57) = array(l57, l57a)
' FadeLights.Callback(57) = "UpdateLamp57"
  FadeLights.obj(58) = array(l58, l58a,f58)
  FadeLights.obj(61) = array(l61)
  FadeLights.obj(62) = array(l62)
  FadeLights.obj(63) = array(l63)
  FadeLights.obj(64) = array(l64)
  FadeLights.obj(65) = array(l65)
  FadeLights.obj(66) = array(l66, l66a)
  FadeLights.obj(67) = array(l67, l67a)
  FadeLights.obj(68) = array(l68, l68a)
  FadeLights.obj(71) = array(f71)
  FadeLights.obj(72) = array(f72)
  FadeLights.obj(73) = array(f73)
  FadeLights.obj(74) = array(f74)
' FadeLights.Callback(72) = "UpdateLamp72"
' FadeLights.Callback(73) = "UpdateLamp73"
' FadeLights.Callback(74) = "UpdateLamp74"
' FadeLights.Callback(75) = "UpdateLamp75" 'Game Saucer
' FadeLights.Callback(76) = "UpdateLamp76" 'Mega Ramp
  FadeLights.obj(77) = array(l77, l77a)
  FadeLights.obj(78) = array(l78, l78a)
  FadeLights.obj(81) = array(l81, l81a)
  FadeLights.obj(82) = array(l82, l82a)
  FadeLights.obj(83) = array(l83, l83a)
  FadeLights.obj(84) = array(l84, l84a)
  FadeLights.obj(85) = array(l85, l85a)
  FadeLights.obj(86) = array(l86, l86a)
' Start WJR Comment
' FadeLights.obj(101) = array(flash101) 'right eye flash
' FadeLights.obj(102) = array(flash102) 'center visor flash
' FadeLights.obj(103) = array(Flasher1R,Flasher1R1,FLasher1R2,Flash103wall) 'lower right flasher
' FadeLights.obj(104) = array(flash104) 'left eye flash
''  FadeLights.obj(105) = array(Flasher2TL,Flasher2TL1,Flasher2TL2,Flasher2TL3,Flasher2TR,Flasher2TR1,Flasher2TR2,Flasher2TR3)
' 'FadeLights.obj(106) = array(Flasher1L,Flasher1L1,FLasher1L2,Flasher1R,Flasher1R1,FLasher1R2)
' FadeLights.obj(106) = array(Flasher1L,Flasher1L1,FLasher1L2,Flash106wall) 'lower left flasher
' FadeLights.obj(107) = array(l107, l107a, l107z,BumpFlash,FlasherB,Flash107wall)
' FadeLights.obj(108) = array(Flasher2L,Flasher2L1,Flasher2L2,Flash108wall)
' FadeLights.obj(109) = array(l109, l109a, l109z, f109)
' FadeLights.obj(110) = array(flasherb1a,flasherb1b,Flasherplateb1,Flash110wall)
' FadeLights.obj(111) = array(flasherb2a,flasherb2b,Flasherplateb2)
' FadeLights.obj(112) = array(flasherb3a,flasherb3b,Flasherplateb3)
' FadeLights.obj(113) = array(flasherb4a,flasherb4b,Flasherplateb4)
' FadeLights.obj(114) = array(flasherb5a,flasherb5b,Flasherplateb5,Flash114wall)
' End WJR Comment

' FadeLights.Callback(115) = "UpdateGIFade"
' FadeLights.Callback(116) = "UpdateGIFade"
  FadeLights.obj(115) = array(shadowson)
  FadeLights.obj(116) = array(shadowsoff)
End Sub

Sub AllLampsOff
    Dim x
    For x = 0 to 200
        SetLamp x, 0
    Next
End Sub

Sub UpdateLamp71
  if FadeLights.state(71) <> 0 Then
    bulb71.image="rightsignbon"
    'bulb71.blenddisablelighting = 0.85
    bulb71.blenddisablelighting = 1.5
  else
    bulb71.image="rightsignboff"
    bulb71.blenddisablelighting = 0.5
  end if
End Sub

Sub UpdateLamp72
  if FadeLights.state(72) <> 0 Then
    bulb72.image="rightsignbon"
    'bulb72.blenddisablelighting = 0.85
    bulb72.blenddisablelighting = 1.5
  else
    bulb72.image="rightsignboff"
    bulb72.blenddisablelighting = 0.5
  end if
End Sub

Sub UpdateLamp73
  if FadeLights.state(73) <> 0 Then
    bulb73.image="rightsignbon"
    'bulb73.blenddisablelighting = 0.85
    bulb73.blenddisablelighting = 1.5
  else
    bulb73.image="rightsignboff"
    bulb73.blenddisablelighting = 0.5
  end if
End Sub

Sub UpdateLamp74
  if FadeLights.state(74) <> 0 Then
    bulb74.image="rightsignbon"
    'bulb74.blenddisablelighting = 0.85
    bulb74.blenddisablelighting = 1.5
  else
    bulb74.image="rightsignboff"
    bulb74.blenddisablelighting = 0.5
  end if
End Sub

Sub UpdateLamp75
  if FadeLights.state(75) <> 0 Then
    rampsigngreen.image="rampsignon"
    'rampsigngreen.blenddisablelighting=1
    rampsigngreen.blenddisablelighting=1.5
    if FadeLights.state(76) <> 0 Then
      rampsign.image="rampsignallon"
    else
      rampsign.image="rampsigngreen"
    end if
  else
    rampsigngreen.image="rampsignoff"
    rampsigngreen.blenddisablelighting=.5
    if FadeLights.state(76) <> 0 Then
      rampsign.image="rampsignyellow"
    else
      rampsign.image="rampsign"
    end if
  end if
End Sub

Sub UpdateLamp76
  if FadeLights.state(76) <> 0 Then
    rampsignyellow.image="rampsignon"
    'rampsignyellow.blenddisablelighting=1
    rampsignyellow.blenddisablelighting=1.5
    if FadeLights.state(75) <> 0 Then
      rampsign.image="rampsignallon"
    else
      rampsign.image="rampsignyellow"
    end if
  else
    rampsignyellow.image="rampsignoff"
    rampsignyellow.blenddisablelighting=.5
    if FadeLights.state(75) <> 0 Then
      rampsign.image="rampsigngreen"
    else
      rampsign.image="rampsign"
    end if
  end if
End Sub
'
'


Sub SetLamp(nr, value)
  ' If the lamp state is not changing, just exit.
  if FadeLights.state(nr) = value then exit sub

  if nr > 100 and nr < 115 Then
    FadeLights.state(nr) = value/255
  else
    FadeLights.state(nr) = value
  end if
  if nr = 71 then UpdateLamp71
  if nr = 72 then UpdateLamp72
  if nr = 73 then UpdateLamp73
  if nr = 74 then UpdateLamp74
  if nr = 75 then UpdateLamp75
  if nr = 76 then UpdateLamp76

' Start WJR Comment
' if nr = 101 then UpdateLamp101
' if nr = 102 then UpdateLamp102
' if nr = 103 then UpdateLamp103
' if nr = 104 then UpdateLamp104
''  if nr = 105 then UpdateLamp105
' if nr = 106 then UpdateLamp106
''  if nr = 107 then UpdateLamp107
' if nr = 108 then UpdateLamp108
' if nr = 110 then UpdateLamp110
' if nr = 111 then UpdateLamp111
' if nr = 112 then UpdateLamp112
' if nr = 113 then UpdateLamp113
' if nr = 114 then UpdateLamp114
' End WJR Comment

End Sub



' *** NFozzy's lamp fade routines ***


Class NullFadingObject : Public Property Let IntensityScale(input) : : End Property : End Class 'todo do better

Class LampFader
  Public FadeSpeedDown(140), FadeSpeedUp(140)
  Private Lock(140), Loaded(140), OnOff(140)
  Public UseFunction
  Private cFilter
  Private UseCallback(140), cCallback(140)
  Public Lvl(140), Obj(140)

  Sub Class_Initialize()
    dim x : for x = 0 to uBound(OnOff)  'Set up fade speeds
      if FadeSpeedDown(x) <= 0 then FadeSpeedDown(x) = 1/100  'fade speed down
      if FadeSpeedUp(x) <= 0 then FadeSpeedUp(x) = 1/80'Fade speed up
      UseFunction = False
      lvl(x) = 0
      OnOff(x) = False
      Lock(x) = True : Loaded(x) = False
    Next

    for x = 0 to uBound(OnOff)    'clear out empty obj
      if IsEmpty(obj(x) ) then Set Obj(x) = NullFader' : Loaded(x) = True
    Next
  End Sub

  Public Property Get Locked(idx) : Locked = Lock(idx) : End Property   'debug.print Lampz.Locked(100)  'debug
  Public Property Get state(idx) : state = OnOff(idx) : end Property
  Public Property Let Filter(String) : Set cFilter = GetRef(String) : UseFunction = True : End Property
  Public Property Let Callback(idx, String) : cCallback(idx) = String : UseCallBack(idx) = True : End Property

  Public Property Let state(ByVal idx, input) 'Major update path
    input = cBool(input)
    if OnOff(idx) = Input then : Exit Property : End If 'discard redundant updates
    OnOff(idx) = input
    Lock(idx) = False
    Loaded(idx) = False
  End Property

  Public Sub TurnOnStates() 'If obj contains any light objects, set their states to 1 (Fading is our job!)
    dim debugstr
    dim idx : for idx = 0 to uBound(obj)
      if IsArray(obj(idx)) then
        dim x, tmp : tmp = obj(idx) 'set tmp to array in order to access it
        for x = 0 to uBound(tmp)
          if typename(tmp(x)) = "Light" then DisableState tmp(x) : debugstr = debugstr & tmp(x).name & " state'd" & vbnewline

        Next
      Else
        if typename(obj(idx)) = "Light" then DisableState obj(idx) : debugstr = debugstr & obj(idx).name & " state'd (not array)" & vbnewline

      end if
    Next
    'debug.print debugstr
  End Sub
  Private Sub DisableState(ByRef aObj) : aObj.FadeSpeedUp = 1000 : aObj.State = 1 : End Sub 'turn state to 1

  Public Sub Update1()   'Handle all boolean numeric fading. If done fading, Lock(True). Update on a '1' interval Timer!
    dim x : for x = 0 to uBound(OnOff)
      if not Lock(x) then 'and not Loaded(x) then
        if OnOff(x) then 'Fade Up
          Lvl(x) = Lvl(x) + FadeSpeedUp(x)
          if Lvl(x) > 1 then Lvl(x) = 1 : Lock(x) = True
        elseif Not OnOff(x) then 'fade down
          Lvl(x) = Lvl(x) - FadeSpeedDown(x)
          if Lvl(x) < 0 then Lvl(x) = 0 : Lock(x) = True
        end if
      end if
    Next
  End Sub


  Public Sub Update() 'Handle object updates. Update on a -1 Timer! If done fading, loaded(x) = True
    dim x,xx : for x = 0 to uBound(OnOff)
      if not Loaded(x) then
        if IsArray(obj(x) ) Then  'if array
          If UseFunction then
            for each xx in obj(x) : xx.IntensityScale = cFilter(Lvl(x)) : Next
          Else
            for each xx in obj(x) : xx.IntensityScale = Lvl(x) : Next
          End If
        else            'if single lamp or flasher
          If UseFunction then
            obj(x).Intensityscale = cFilter(Lvl(x))
          Else
            obj(x).Intensityscale = Lvl(x)
          End If
        end if
        ' Sleazy hack for regional decimal point problem
        If UseCallBack(x) then execute cCallback(x) & " CSng(" & CInt(10000 * Lvl(x)) & " / 10000)" 'Callback
        If Lock(x) Then
          if Lvl(x) = 1 or Lvl(x) = 0 then Loaded(x) = True 'finished fading
        end if
      end if
    Next
  End Sub
End Class


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

Sub PlaySoundAtBall(soundname)
    PlaySoundAt soundname, ActiveBall
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

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
    Vol = Csng(BallVel(ball) ^2 / 2000)
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

Const tnob = 5 ' total number of balls
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
    StopSound("fx_Rolling_Wood" & b)
    StopSound("fx_Rolling_Plastic" & b)
    StopSound("fx_Rolling_Metal" & b)
  Next

  ' exit the sub if no balls on the table
  If UBound(BOT) = -1 Then Exit Sub

  ' play the rolling sound for each ball
  For b = 0 to UBound(BOT)
    If BallVel(BOT(b) ) > 1 Then

      ' ***Ball on WOOD playfield***
      if BOT(b).z < 27 Then
        PlaySound("fx_Rolling_Wood" & b), -1, Vol(BOT(b) )/2, AudioPan(BOT(b) )/5, 0, Pitch(BOT(b) ), 1, 0, AudioFade(BOT(b) )
        rolling(b) = True
        StopSound("fx_Rolling_Plastic" & b)
        StopSound("fx_Rolling_Metal" & b)
      ' ***Ball on METAL ramp***
      ElseIf BOT(b).z > 90 and InRect(BOT(b).x, BOT(b).y, 700, 1043, 765,1043, 765,1317, 700,1317) Then
        PlaySound("fx_Rolling_Metal" & b), -1, Vol(BOT(b) ), AudioPan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0, AudioFade(BOT(b) )
        rolling(b) = True
        StopSound("fx_Rolling_Wood" & b)
        StopSound("fx_Rolling_Plastic" & b)
      ' ***Ball on PLASTIC ramp***
      Elseif  BOT(b).z > 100 and InRect(BOT(b).x, BOT(b).y, 590,395,940,395,940, 951, 670, 1095) Then
        PlaySound("fx_Rolling_Plastic" & b), -1, Vol(BOT(b) ), AudioPan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0, AudioFade(BOT(b) )
        rolling(b) = True
        StopSound("fx_Rolling_Wood" & b)
        StopSound("fx_Rolling_Metal" & b)
      Elseif  BOT(b).z > 30 and InRect(BOT(b).x, BOT(b).y, 43, 523, 130, 500,213, 745, 118, 780) Then
        PlaySound("fx_Rolling_Plastic" & b), -1, Vol(BOT(b) ), AudioPan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0, AudioFade(BOT(b) )
        rolling(b) = True
        StopSound("fx_Rolling_Wood" & b)
        StopSound("fx_Rolling_Metal" & b)
      Elseif  BOT(b).z > 90 and InRect(BOT(b).x, BOT(b).y, 0,0,928,0,928,400,0,520) Then
        PlaySound("fx_Rolling_Plastic" & b), -1, Vol(BOT(b) ),AudioPan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0, AudioFade(BOT(b) )
        rolling(b) = True
        StopSound("fx_Rolling_Wood" & b)
        StopSound("fx_Rolling_Metal" & b)
      Elseif  BOT(b).z > 30 and InRect(BOT(b).x, BOT(b).y, 566,51,853,45,861,381,573,384) Then
        PlaySound("fx_Rolling_Plastic" & b), -1, Vol(BOT(b) ), AudioPan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0, AudioFade(BOT(b) )
        rolling(b) = True
        StopSound("fx_Rolling_Wood" & b)
        StopSound("fx_Rolling_Metal" & b)
      Else
        if rolling(b) = true Then
          rolling(b) = False
          StopSound("fx_Rolling_Wood" & b)
          StopSound("fx_Rolling_Plastic" & b)
          StopSound("fx_Rolling_Metal" & b)
        end if
      End If
    Else
      If rolling(b) = True Then
        StopSound("fx_Rolling_Wood" & b)
        StopSound("fx_Rolling_Plastic" & b)
        StopSound("fx_Rolling_Metal" & b)
        rolling(b) = False
      End If

      if BOT(b).z > 120 and InRect(BOT(b).x, BOT(b).y, 634,234,690,206,711,247,669,295) Then
        if BOT(b).vely < 0 and ballvel(BOT(b)) < 1 Then
          BOT(b).vely = BOT(b).vely * 2
          if BOT(b).velx > 0 then BOT(b).velx = -0.1
          'debug.print ballvel(BOT(b))
        end if
      elseif BOT(b).z > 120 and InRect(BOT(b).x, BOT(b).y, 645,162,708,172,690,206,634,234) Then
        if BOT(b).vely < 0 and ballvel(BOT(b)) < 1 Then
          BOT(b).vely = BOT(b).vely * 2
          if BOT(b).velx > 0 then BOT(b).velx = -0.1
        end if
      elseif BOT(b).z > 120 and InRect(BOT(b).x, BOT(b).y, 697,120,744,164,708,172,645,162) Then
        if BOT(b).velx > 0 and ballvel(BOT(b)) < 1 Then
          BOT(b).velx = BOT(b).velx * 2
          if BOT(b).vely > 0 then BOT(b).vely = -0.1
        end if
      end if
    End If

    '***Ball Drop Sounds***

    If BOT(b).VelZ < -1 and BOT(b).z < 50 and BOT(b).z > 27 Then 'height adjust for ball drop sounds
      if not InRect(BOT(b).x, BOT(b).y, 625,55,825,55,825,155,550,225) Then
        PlaySoundAtBOTBallZ "fx_ball_drop" & b, BOT(b)
        'debug.print "Drop" & BOT(b).velz
      end if
    elseif ballvel(BOT(b)) > 10 and InRect(BOT(b).x, BOT(b).y, 690,170,740,170,740,220,690,220) Then
        PlaySoundAtBOTBallZ "fx_glass", BOT(b)
    End If
  Next
End Sub

Sub PlaySoundAtBOTBall(sound, BOT)
    PlaySound sound, 0, Vol(BOT), AudioPan(BOT), 0, Pitch(BOT), 0, 1, AudioFade(BOT)
End Sub

Sub PlaySoundAtBOTBallZ(sound, BOT)
    PlaySound sound, 0, ABS(BOT.velz)/17, AudioPan(BOT), 0, Pitch(BOT), 1, 0, AudioFade(BOT)
End Sub

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


'**********************
' Ball Collision Sound
'**********************

Sub OnBallBallCollision(ball1, ball2, velocity)
  PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 2000, AudioPan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub


'*****************************************
' ninuzzu's FLIPPER SHADOWS
'*****************************************

sub FlipperUpdate()
  batleftshadow.rotz = LeftFlipper.CurrentAngle
  batrightshadow.rotz  = RightFlipper.CurrentAngle
  plungegate_prim.RotX = Gate4.CurrentAngle + 90
  corkwire.RotX = gate5.CurrentAngle
End Sub

'*****************************************
' ninuzzu's BALL SHADOW
'*****************************************
Dim BallShadow
BallShadow = Array (BallShadow1,BallShadow2,BallShadow3,BallShadow4,BallShadow5)


Sub BallShadowUpdate()
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
            BallShadow(b).X = ((BOT(b).X) - (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/21)) + 6
        Else
            BallShadow(b).X = ((BOT(b).X) + (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/21)) - 6
        End If
        ballShadow(b).Y = BOT(b).Y + 4
        If BOT(b).Z > 20 Then
            BallShadow(b).visible = 1
        Else
            BallShadow(b).visible = 0
        End If
    Next
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

Sub Table1_MusicDone()
End Sub

' **** Replacement mech class that moves smoothly.   Carbon copy of the cvpmMech
Class cvpmMyMech
  Public Sol1, Sol2, MType, Length, Steps, Acc, Ret
  Private mMechNo, mNextSw, mSw(), mLastPos, mLastSpeed, mCallback

  Private Sub Class_Initialize
    ReDim mSw(10)
    gNextMechNo = gNextMechNo + 1 : mMechNo = gNextMechNo : mNextSw = 0 : mLastPos = 0 : mLastSpeed = 0
    MType = 0 : Length = 0 : Steps = 0 : Acc = 0 : Ret = 0 : vpmTimer.addResetObj Me
  End Sub

  Public Sub AddSw(aSwNo, aStart, aEnd)
    mSw(mNextSw) = Array(aSwNo, aStart, aEnd, 0)
    mNextSw = mNextSw + 1
  End Sub

  Public Sub AddPulseSwNew(aSwNo, aInterval, aStart, aEnd)
    If Controller.Version >= "01200000" Then
      mSw(mNextSw) = Array(aSwNo, aStart, aEnd, aInterval)
    Else
      mSw(mNextSw) = Array(aSwNo, -aInterval, aEnd - aStart + 1, 0)
    End If
    mNextSw = mNextSw + 1
  End Sub

  Public Sub Start
    Dim sw, ii
    With Controller
      .Mech(1) = Sol1 : .Mech(2) = Sol2 : .Mech(3) = Length
      .Mech(4) = Steps : .Mech(5) = MType : .Mech(6) = Acc : .Mech(7) = Ret
      ii = 10
      For Each sw In mSw
        If IsArray(sw) Then
          .Mech(ii) = sw(0) : .Mech(ii+1) = sw(1)
          .Mech(ii+2) = sw(2) : .Mech(ii+3) = sw(3)
          ii = ii + 10
        End If
      Next
      .Mech(0) = mMechNo
    End With
    If IsObject(mCallback) Then mCallBack 0, 0, 0 : mLastPos = 0 : vpmTimer.EnableUpdate Me, True, True  ' <------- All for this.
  End Sub

  Public Property Get Position : Position = Controller.GetMech(mMechNo) : End Property
  Public Property Get Speed    : Speed = Controller.GetMech(-mMechNo)   : End Property
  Public Property Let Callback(aCallBack) : Set mCallback = aCallBack : End Property

  Public Sub Update
    Dim currPos, speed
    currPos = Controller.GetMech(mMechNo)
    speed = Controller.GetMech(-mMechNo)
    If currPos < 0 Or (mLastPos = currPos And mLastSpeed = speed) Then Exit Sub
    mCallBack currPos, speed, mLastPos : mLastPos = currPos : mLastSpeed = speed
  End Sub

  Public Sub Reset : Start : End Sub

End Class

Dim EnableBallControl
EnableBallControl = false 'Change to true to enable manual ball control (or press C in-game) via the arrow keys and B (boost movement) keys

'*****************************************
'   rothbauerw's Manual Ball Control
'*****************************************

Dim BCup, BCdown, BCleft, BCright
Dim ControlBallInPlay, ControlActiveBall
Dim BCvel, BCyveloffset, BCboostmulti, BCboost

BCboost = 1       'Do Not Change - default setting
BCvel = 4       'Controls the speed of the ball movement
BCyveloffset = -0.01  'Offsets the force of gravity to keep the ball from drifting vertically on the table, should be negative
BCboostmulti = 3    'Boost multiplier to ball veloctiy (toggled with the B key)

ControlBallInPlay = false

Sub StartBallControl_Hit()
  Set ControlActiveBall = ActiveBall
  ControlBallInPlay = true
End Sub

Sub StopBallControl_Hit()
  ControlBallInPlay = false
End Sub

Sub BallControlTimer_Timer()
  If EnableBallControl and ControlBallInPlay then
    If BCright = 1 Then
      ControlActiveBall.velx =  BCvel*BCboost
    ElseIf BCleft = 1 Then
      ControlActiveBall.velx = -BCvel*BCboost
    Else
      ControlActiveBall.velx = 0
    End If

    If BCup = 1 Then
      ControlActiveBall.vely = -BCvel*BCboost
    ElseIf BCdown = 1 Then
      ControlActiveBall.vely =  BCvel*BCboost
    Else
      ControlActiveBall.vely = bcyveloffset
    End If
  End If
End Sub

Option Explicit
Randomize

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Const cGameName="badgirls",UseSolenoids=1,UseLamps=0,UseGI=0,SSolenoidOn="SolOn",SSolenoidOff="SolOff", SCoin="coin"

LoadVPM "01120100","sys80.vbs",3.02

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
 SolCallback(1) = "Solenoid1"         'sol1 3 position 7 bank
 SolCallback(2) = "Solenoid2"         'sol2 left upkicker
 SolCallback(3) = "Solenoid3"         'sol3 Large Hat Lights
 SolCallback(4) = "Solenoid4"         'sol4 VariReset
 SolCallback(5) = "Solenoid5"         'sol5 4 position 7 bank
 SolCallback(6) = "Solenoid6"         'sol6 right upkicker
 SolCallback(7) = "Solenoid7"         'sol7 EightBall Drop
 SolCallback(8) =  "vpmSolSound SoundFX(""Knocker"",DOFKnocker),"
 SolCallback(9) = "bsTrough.SolIn"            'sol.9 Outhole

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
     If Enabled Then
         PlaySoundAtVol SoundFX("fx_Flipperup",DOFContactors), Leftflipper, 1:LeftFlipper.RotateToEnd
     Else
         PlaySoundatVol SoundFX("fx_Flipperdown",DOFContactors), Leftflipper, 1:LeftFlipper.RotateToStart
     End If
  End Sub

Sub SolRFlipper(Enabled)
     If Enabled Then
         PlaySoundAtVol SoundFX("fx_Flipperup",DOFContactors), RightFlipper, 1:RightFlipper.RotateToEnd:RightFlipper1.RotateToEnd
     Else
         PlaySoundAtVol SoundFX("fx_Flipperdown",DOFContactors), RightFlipper, 1:RightFlipper.RotateToStart:RightFlipper1.RotateToStart
     End If
End Sub
'**********************************************************************************************************

'Solenoid Controlled toys
'**********************************************************************************************************

Sub solTrough(Enabled) 'controlled by light ID 2
  If Enabled Then
    bsTrough.ExitSol_On
  End If
 End Sub

 Sub Solenoid1(Enabled)
  If Controller.Lamp(12) Then 'S Relay enable
    SetLamp 151,Enabled 'Upper Left Flahsers X2
  Else
    dtL.SolunHit 1,Enabled:
    dtL.SolunHit 2,Enabled:
    dtL.SolunHit 3,Enabled:
  End If
 End Sub

 Sub Solenoid2(Enabled)
  If Controller.Lamp(12) Then 'S Relay enable
    SetLamp 152, Enabled  ' Flasher Drop Targets X2
  Else
    if Controller.switch (45) then
      PlaySoundAtVol SoundFX("Popper",DOFContactors), TopVuk, 1
      TopVUK.DestroyBall
      Set raiseball = TopVUK.CreateBall
      raiseballsw = True
      TopVukraiseballtimer.Enabled = True
      TopVUK.Enabled=TRUE
      Controller.switch (45) = False
    else
      PlaySoundAtVol "Popper", TopVuk, 1
    end if
  End If
 End Sub

 Sub Solenoid3(Enabled)
  If Controller.Lamp(12) Then 'S Relay enable
    SetLamp 153,Enabled  'Flasher Right Drain
  Else
    SetLamp 150,Enabled 'red box slingshots X2 + PF
  End If
 End Sub

 Sub Solenoid4(Enabled)
  If Controller.Lamp(12) Then 'S Relay enable
    SetLamp 154,Enabled  'Flasher Left Drain
  Else
    If Enabled Then SolVariReset
    End If
 End Sub

 Sub Solenoid5(Enabled)
  If Controller.Lamp(12) Then 'S Relay enable
    SetLamp 155,Enabled ' red box upper right
  Else
    dtL.SolunHit 4,Enabled:
    dtL.SolunHit 5,Enabled:
    dtL.SolunHit 6,Enabled:
    dtL.SolunHit 7,Enabled:
  End If
 End Sub

 Sub Solenoid6(Enabled)
  If Controller.Lamp(12) Then 'S Relay enable
    SetLamp 156,Enabled'Dome Flasher X2
  Else
    if Controller.switch (55) then
      PlaySoundAtVol SoundFX("Popper",DOFContactors), MidVukTop, 1
      MidVUKTop.DestroyBall
      Set dropball = MidVUKTop.CreateBall
      dropballsw = True
      MidVukdropballtimer.Enabled = True
      MidVUKTop.Enabled=TRUE
      Controller.switch (55) = False
    else
      PlaySoundAtVol "Popper", MidVukTop, 1
    end if
  End If
 End Sub

 Sub Solenoid7(Enabled)
  If Controller.Lamp(12) Then 'S Relay enable
    SetLamp 157,Enabled 'Plunger Lane Flasher X2
  Else
    dtT.SolunHit 1,Enabled:
  End If
 End Sub

'Playfield GI
Sub PFGI(Enabled)
  If Enabled Then
    dim xx
    For each xx in GI:xx.State = 0: Next
        PlaySound "fx_relay"

  Else
    For each xx in GI:xx.State = 1: Next
        PlaySound "fx_relay"


  End If
End Sub

'Left Sequential GI Lights controlled by relay B
Sub RelayB(Enabled)
  If Enabled Then
    Timer1.Enabled = 0
  Else
        PlaySound "fx_relay"
  End If
End Sub

Dim AuxRampLamps1, AuxLampCount1

Sub Timer1_Timer
  AuxLampCount1 = (AuxLampCount1 + 1) Mod 9
  Select Case AuxLampCount
    Case 0:
       SetLamp 101, 1: SetLamp 102, 0: SetLamp 103, 0 : SetLamp 104, 0 : SetLamp 105, 0
    Case 1:
       SetLamp 101, 0: SetLamp 102, 1: SetLamp 103, 0 : SetLamp 104, 0 : SetLamp 105, 0
    Case 3:
       SetLamp 101, 0: SetLamp 102, 0: SetLamp 103, 1 : SetLamp 104, 0 : SetLamp 105, 0
    Case 4:
       SetLamp 101, 0: SetLamp 102, 0: SetLamp 103, 0 : SetLamp 104, 1 : SetLamp 105, 0
    Case 5:
       SetLamp 101, 0: SetLamp 102, 0: SetLamp 103, 0 : SetLamp 104, 0 : SetLamp 105, 1
    Case 6:
       SetLamp 101, 0: SetLamp 102, 0: SetLamp 103, 0 : SetLamp 104, 1 : SetLamp 105, 0
    Case 7:
       SetLamp 101, 0: SetLamp 102, 0: SetLamp 103, 1 : SetLamp 104, 0 : SetLamp 105, 0
    Case 8:
       SetLamp 101, 0: SetLamp 102, 1: SetLamp 103, 0 : SetLamp 104, 0 : SetLamp 105, 0
  End Select
End Sub


'Backwall Lights controlled by relay Q
Sub RelayQ(Enabled)
  If Enabled Then
    Timer2.Enabled = 1
    dim xxx
    For each xxx in BGI:xxx.visible = 1: Next
  Else
    Timer2.Enabled = 0
    For each xxx in BGI:xxx.visible = 0: Next
  End If
End Sub

Dim AuxRampLamps, AuxLampCount

Sub Timer2_Timer
  AuxLampCount = (AuxLampCount + 1) Mod 11
  Select Case AuxLampCount
    Case 0:
       SetLamp 111, 1: SetLamp 112, 0: SetLamp 113, 0 : SetLamp 114, 0 : SetLamp 115, 0 : SetLamp 116, 0 : SetLamp 117, 0 : SetLamp 118, 0 : SetLamp 119, 0 : SetLamp 120, 0
    Case 1:
       SetLamp 111, 0: SetLamp 112, 1: SetLamp 113, 0 : SetLamp 114, 0 : SetLamp 115, 0 : SetLamp 116, 0 : SetLamp 117, 0 : SetLamp 118, 0 : SetLamp 119, 0 : SetLamp 120, 0
    Case 3:
       SetLamp 111, 0: SetLamp 112, 0: SetLamp 113, 1 : SetLamp 114, 0 : SetLamp 115, 0 : SetLamp 116, 0 : SetLamp 117, 0 : SetLamp 118, 0 : SetLamp 119, 0 : SetLamp 120, 0
    Case 4:
       SetLamp 111, 0: SetLamp 112, 0: SetLamp 113, 0 : SetLamp 114, 1 : SetLamp 115, 0 : SetLamp 116, 0 : SetLamp 117, 0 : SetLamp 118, 0 : SetLamp 119, 0 : SetLamp 120, 0
    Case 5:
       SetLamp 111, 0: SetLamp 112, 0: SetLamp 113, 0 : SetLamp 114, 0 : SetLamp 115, 1 : SetLamp 116, 0 : SetLamp 117, 0 : SetLamp 118, 0 : SetLamp 119, 0 : SetLamp 120, 0
    Case 6:
       SetLamp 111, 0: SetLamp 112, 0: SetLamp 113, 0 : SetLamp 114, 0 : SetLamp 115, 0 : SetLamp 116, 1 : SetLamp 117, 0 : SetLamp 118, 0 : SetLamp 119, 0 : SetLamp 120, 0
    Case 7:
       SetLamp 111, 0: SetLamp 112, 0: SetLamp 113, 0 : SetLamp 114, 0 : SetLamp 115, 0 : SetLamp 116, 0 : SetLamp 117, 1 : SetLamp 118, 0 : SetLamp 119, 0 : SetLamp 120, 0
    Case 8:
       SetLamp 111, 0: SetLamp 112, 0: SetLamp 113, 0 : SetLamp 114, 0 : SetLamp 115, 0 : SetLamp 116, 0 : SetLamp 117, 0 : SetLamp 118, 1 : SetLamp 119, 0 : SetLamp 120, 0
    Case 9:
       SetLamp 111, 0: SetLamp 112, 0: SetLamp 113, 0 : SetLamp 114, 0 : SetLamp 115, 0 : SetLamp 116, 0 : SetLamp 117, 0 : SetLamp 118, 0 : SetLamp 119, 1 : SetLamp 120, 0
    Case 10:
       SetLamp 111, 0: SetLamp 112, 0: SetLamp 113, 0 : SetLamp 114, 0 : SetLamp 115, 0 : SetLamp 116, 0 : SetLamp 117, 0 : SetLamp 118, 0 : SetLamp 119, 0 : SetLamp 120, 1
  End Select
End Sub


'**********************************************************************************************************

'Initiate Table
'**********************************************************************************************************

 Dim bsTrough, bsLeftSaucer, bsRightSaucer, dtT, dtL

Sub Table1_Init
  vpmInit Me
  On Error Resume Next
    With Controller
    .GameName = cGameName
    If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
    .SplashInfoLine = "Bad Girls Gottlieb/Premier 1990"&chr(13)&"You Suck"
    .HandleMechanics=0
    .HandleKeyboard=0
    .ShowDMDOnly=1
    .ShowFrame=0
    .ShowTitle=0
        .hidden = 1
         On Error Resume Next
         .Run GetPlayerHWnd
         If Err Then MsgBox Err.Description
         On Error Goto 0
     End With
     On Error Goto 0

  PinMAMETimer.Interval=PinMAMEInterval
  PinMAMETimer.Enabled=1
  vpmNudge.TiltSwitch = 57
  vpmNudge.Sensitivity = 2
  vpmNudge.TiltObj = Array(Bumper1,Bumper2,Bumper3,LeftSlingshot,RightSlingshot)

  Set bsTrough = New cvpmBallStack
    bsTrough.InitSw 65, 76, 75, 0, 0, 0, 0, 0
    bsTrough.InitKick BallRelease, 90, 15
    bsTrough.InitExitSnd SoundFX("ballrelease",DOFContactors), SoundFX("Solenoid",DOFContactors)
    bsTrough.Balls = 2

  Set dtT=New cvpmDropTarget
    dtT.InitDrop sw70,70
    dtT.InitSnd SoundFX("DTDrop",DOFContactors),SoundFX("DTReset",DOFContactors)

  Set dtL=New cvpmDropTarget
    dtL.InitDrop Array(sw0,sw10,sw20,sw30,sw40,sw50,sw60),Array(0,10,20,30,40,50,60)
    dtL.InitSnd SoundFX("DTDrop",DOFContactors),SoundFX("DTReset",DOFContactors)

 End Sub

Sub table1_Paused:Controller.Pause = 1:End Sub
Sub table1_unPaused:Controller.Pause = 0:End Sub
Sub table1_exit:Controller.stop:End Sub

'**********************************************************************************************************
'Plunger code
'**********************************************************************************************************

Sub Table1_KeyDown(ByVal KeyCode)
  If keycode = PlungerKey Then Plunger.Pullback:PlaySoundAtVol "plungerpull", Plunger, 1
  If keycode = RightFlipperkey Then Controller.Switch(15) = 1

  If KeyDownHandler(keycode) Then Exit Sub
End Sub

Sub Table1_KeyUp(ByVal KeyCode)
  If keycode = PlungerKey Then Plunger.Fire:PlaySoundAtVol "plunger", Plunger, 1
  If keycode = RightFlipperkey Then Controller.Switch(15) = 0

  If KeyUpHandler(keycode) Then Exit Sub
End Sub

'**********************************************************************************************************

 ' Drain hole and kickers
Sub Drain_Hit:bsTrough.addball me : PlaySoundAtVol "drain", Drain, 1 : End Sub

 '***********************************
 'Top Left Raising VUK
 '***********************************
 Dim raiseballsw, raiseball

 Sub TopVUK_Hit()
  TopVUK.Enabled=FALSE
  Controller.switch (45) = True
  PlaySoundAtVol "popper_ball", ActiveBall, 1
 End Sub

 Sub TopVukraiseballtimer_Timer()
  If raiseballsw = True then
    raiseball.z = raiseball.z + 10
    If raiseball.z > 140 then
      TopVUK.Kick 180, 10
      Set raiseball = Nothing
      TopVukraiseballtimer.Enabled = False
      raiseballsw = False
    End If
  End If
 End Sub

 '********************
 'Middle Raising VUK
 '********************
 Dim dropballsw, dropball

Sub MidVUKTop_Hit()
  MidVUKTop.Enabled=FALSE
  Controller.switch (55) = True
  PlaySoundAtVol "popper_ball", ActiveBall, 1
End Sub

 Sub MidVukdropballtimer_Timer()
  If dropballsw = True then
    dropball.z = dropball.z + 10
    If dropball.z > 100 then
      MidVUKTop.Kick 45, 10
      Set dropball = Nothing
      PlaySoundAtVol "fx_ballrampdrop", MidVukTop, 1
      MidVukdropballtimer.Enabled = False
      dropballsw = False
    End If
  End If
 End Sub

'Drop Targets
 Sub Sw0_Dropped:dtL.Hit 1 :End Sub
 Sub Sw10_Dropped:dtL.Hit 2 :End Sub
 Sub Sw20_Dropped:dtL.Hit 3 :End Sub
 Sub Sw30_Dropped:dtL.Hit 4 :End Sub
 Sub Sw40_Dropped:dtL.Hit 5 :End Sub
 Sub Sw50_Dropped:dtL.Hit 6 :End Sub
 Sub Sw60_Dropped:dtL.Hit 7 :End Sub

 Sub Sw70_Dropped:dtT.Hit 1 :End Sub

 'Stand Up Targets
Sub sw1_hit:vpmTimer.pulseSw 1 : End Sub
Sub sw11_hit:vpmTimer.pulseSw 11 : End Sub
Sub sw21_hit:vpmTimer.pulseSw 21 : End Sub
Sub sw31_hit:vpmTimer.pulseSw 31 : End Sub
Sub sw41_hit:vpmTimer.pulseSw 41 : End Sub
Sub sw51_hit:vpmTimer.pulseSw 51 : End Sub

'Wire Triggers
 Sub sw3_Hit:Controller.Switch(3)=1 : PlaySoundAtVol "rollover", ActiveBall, 1 : End Sub
 Sub sw3_UnHit:Controller.Switch(3)=0:End Sub
 Sub sw13_Hit:Controller.Switch(13)=1 : PlaySoundAtVol "rollover", ActiveBall, 1 : End Sub
 Sub sw13_UnHit:Controller.Switch(13)=0:End Sub
 Sub sw23_Hit:Controller.Switch(23)=1 : PlaySoundAtVol "rollover", ActiveBall, 1 : End Sub
 Sub sw23_UnHit:Controller.Switch(23)=0:End Sub
 Sub sw33_Hit:Controller.Switch(33)=1 : PlaySoundAtVol "rollover", ActiveBall, 1 : End Sub
 Sub sw33_UnHit:Controller.Switch(33)=0:End Sub
 Sub sw43_Hit:Controller.Switch(43)=1 : PlaySoundAtVol "rollover", ActiveBall, 1 : End Sub
 Sub sw43_UnHit:Controller.Switch(43)=0:End Sub
 Sub sw53_Hit:Controller.Switch(53)=1 : PlaySoundAtVol "rollover", ActiveBall, 1 : End Sub
 Sub sw53_UnHit:Controller.Switch(53)=0:End Sub
 Sub sw63_Hit:Controller.Switch(63)=1 : PlaySoundAtVol "rollover", ActiveBall, 1 : End Sub
 Sub sw63_UnHit:Controller.Switch(63)=0:End Sub
 Sub sw73_Hit:Controller.Switch(73)=1 : PlaySoundAtVol "rollover", ActiveBall, 1 : End Sub
 Sub sw73_UnHit:Controller.Switch(73)=0:End Sub

'Star Rollover
 Sub sw72a_Hit:Controller.Switch(72)=1 : PlaySoundAtVol "rollover", ActiveBall, 1 : End Sub
 Sub sw72a_UnHit:Controller.Switch(72)=0:End Sub
 Sub sw72b_Hit:Controller.Switch(72)=1 : PlaySoundAtVol "rollover", ActiveBall, 1 : End Sub
 Sub sw72b_UnHit:Controller.Switch(72)=0:End Sub

'Bumpers
Sub Bumper1_Hit : vpmTimer.PulseSw(52) : PlaySoundAtVol SoundFX("fx_bumper1",DOFContactors), ActiveBall, 1: End Sub
Sub Bumper2_Hit : vpmTimer.PulseSw(42) : PlaySoundAtVol SoundFX("fx_bumper1",DOFContactors), ActiveBall, 1: End Sub
Sub Bumper3_Hit : vpmTimer.PulseSw(62) : PlaySoundatVol SoundFX("fx_bumper1",DOFContactors), ActiveBall, 1: End Sub

'Spinners
Sub sw71_Spin:vpmTimer.PulseSw 71 : PlaySoundAtVol "fx_spinner", sw71, 1 : End Sub
Sub sw61_Spin:vpmTimer.PulseSw 61 : PlaySoundAtVol "fx_spinner", sw61, 1 : End Sub

'Generic Sounds
Sub Trigger1_Hit:PlaySoundAtVol "fx_ballrampdrop", ActiveBall, 1:End Sub
Sub Trigger2_Hit:PlaySoundAtVol "fx_ballrampdrop", ActiveBall, 1:End Sub

Sub Trigger3_Hit:PlaySoundAtVol "Wire Ramp", ActiveBall, 1:End Sub
Sub Trigger4_Hit:PlaySoundAtVol "Wire Ramp", ActiveBall, 1:End Sub

'**********************************************************************************************************
 'VariTarget
'**********************************************************************************************************

 Dim VariResist,VariDir:VariResist = 1.5
 Sub VariTriggers_Hit(idx)
  If ActiveBall.VelY < -1 Then
    If idx > 0 Then VariTriggers(idx-1).enabled = 0
    ActiveBall.VelY = ActiveBall.VelY + (dCos(15)*VariResist)
    ActiveBall.VelX = ActiveBall.VelX - (dSin(15)*VariResist)
    VariDir = 1:VariTimer.enabled = 1
    Vari.rotx = idx + 1
  End If
 End Sub

 Sub VariTriggers_UnHit(idx)
  VariDir = 0
 End Sub

 Sub VariTimer_Timer
   Dim VPos(2)
  Select Case VariDir
    Case 0,1
      VPos(0) = (Vari.rotx - 3)\6 : VPos(1) = (VPos(0)*6) + 3 : VPos(2) = VPos(0)
      If vari.rotx > 2 Then VPos(2) = VPos(0)+1
      Select Case VPos(2)
        Case 1
          Controller.Switch(2)=1

        Case 3
          Controller.Switch(2)=1
          Controller.Switch(12)=1

        Case 5
          Controller.Switch(2)=1
          Controller.Switch(12)=1
          Controller.Switch(22)=1

        Case 6
          Controller.Switch(2)=1
          Controller.Switch(12)=1
          Controller.Switch(22)=1
          Controller.Switch(32)=1

      End Select
      If VariDir = 0 Then
        If Vari.rotx > VPos(1) Then Vari.rotx = Vari.rotx - 1:VariTriggers(vari.rotx).enabled = 1
        If Vari.rotx = VPos(1) Then me.enabled = 0:Exit Sub
      End If
    Case -1
      If VariDir = -1 Then
        If Vari.rotx > 6 Then
          Vari.rotx = Vari.rotx - 7
        ElseIf Vari.rotx > 0 Then
          Vari.rotx = 0:PlaySound "metalhit2"
        End If
        If Vari.rotx = 0 Then VariDir = 0:me.enabled = 0:Exit Sub
      End If
  End Select
 End Sub

 Sub SolVariReset
   Dim x
  Controller.Switch(32)=0
  Controller.Switch(22)=0
  Controller.Switch(12)=0
  Controller.Switch(2)=0

  For each x in VariTriggers:x.enabled = 1:Next
  VariDir = -1:VariTimer.enabled = 1
 End Sub

' Function dsin(angle)
'
' Calculates the sine of an angle given in degrees
'
' dsin = Sin(angle * 3.14159 / 180)
' End Function

Function dcos(angle)
'
' Calculates the cosine of an angle given in degrees
'
dcos = Cos(angle * 3.14159 / 180)
End Function


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

      'Special Handling

      If chgLamp(ii,0) = 0 Then RelayQ chgLamp(ii,1) 'Relay Q Game Over
  '   If chgLamp(ii,0) = 1 Then RelayT chgLamp(ii,1) 'Relay T Tilt
      If chgLamp(ii,0) = 2 Then solTrough chgLamp(ii,1) 'Ballrelease
  '   If chgLamp(ii,0) = 12 Then RelayS chgLamp(ii,1) 'Relay S Switching
  '   If chgLamp(ii,0) = 13 Then RelayA chgLamp(ii,1) 'Relay A Trip
      If chgLamp(ii,0) = 14 Then RelayB chgLamp(ii,1) 'Relay B Rerack
  '   If chgLamp(ii,0) = 15 Then RelayC chgLamp(ii,1) 'Relay C Solids / Stripes

        Next
    End If
    UpdateLamps
End Sub


Sub UpdateLamps()


NFadeLm 3, L3
NFadeL 3, L3a
NFadeLm 4, L4
NFadeL 5, L5
NFadeL 6, L6
NFadeL 7, L7
NFadeL 8, L8
NFadeL 9, L9
NFadeL 10, L10
NFadeL 11, L11

NFadeL 16, L16
NFadeL 17, L17
NFadeLm 18, L18
NFadeL 18, L18a
NFadeLm 19, L19
NFadeL 19, L19a
NFadeL 20, L20
NFadeL 21, L21
NFadeL 22, L22
NFadeLm 23, L23
NFadeL 23, L23a
NFadeLm 24, L24
NFadeL 24, L24a
NFadeL 25, L25
NFadeLm 26, L26
NFadeL 26, L26a
NFadeLm 27, L27
NFadeL 27, L27a
NFadeL 28, L28
NFadeLm 29, L29a
NFadeL 29, L29a
NFadeLm 30, L30
NFadeLm 30, L30a
NFadeL 30, L30b
NFadeLm 31, L31
NFadeLm 31, L31a
NFadeL 31, L31b
NFadeLm 32, L32
NFadeLm 32, L32a
NFadeL 32, L32b
NFadeLm 33, L33a
NFadeLm 33, L33b
NFadeL 33, L33
NFadeLm 34, L34a
NFadeLm 34, L34b
NFadeL 34, L34
NFadeLm 35, L35
NFadeL 35, L35a
NFadeLm 36, L36
NFadeL 36, L36a
NFadeLm 37, L37
NFadeL 37, L37a
NFadeLm 38, L38
NFadeL 38, L38a
NFadeLm 39, L39
NFadeL 39, L39a
NFadeLm 40, L40
NFadeL 40, L40a
NFadeLm 41, L41
NFadeL 41, L41a
NFadeLm 42, L42
NFadeL 42, L42a
NFadeLm 43, L43
NFadeL 43, L43a
NFadeLm 44, L44
NFadeL 44, L44a
NFadeL 45, L45
NFadeL 46, L46
NFadeL 47, L47

'Solenoid Controlled Flasheers

  NFadeLm 151, L151a
  NFadeLm 151, L151b
  NFadeLm 151, L151c
  NFadeL 151, L151d

  NFadeLm 152, L152a
  NFadeLm 152, L152b
  NFadeLm 152, L152c
  NFadeL 152, L152d

  NFadeLm 153, L153a
  NFadeL 153, L153b

    NFadeObjm 150, P150a, "red-gradON", "red-grad"
    NFadeObjm 150, P150b, "red-gradON", "red-grad"
  NFadeLm 150, L150a
  NFadeLm 150, L150b
  NFadeLm 150, L150c
  NFadeL 150, L150d

  NFadeLm 154, L154a
  NFadeL 154, L154b

    NFadeObjm 155, P155, "red-gradON", "red-grad"
  NFadeL 155, L155a

    NFadeObjm 156, P156a, "dome2_0_clearON", "dome2_0_clear"   ' Right Dome
    NFadeObjm 156, P156b, "dome2_0_clearON", "dome2_0_clear"   ' Right Dome
  NFadeLm 156, L156a
  NFadeL 156, L156b

  NFadeLm 157, L157a
  NFadeLm 157, L157b
  NFadeLm 157, L157c
  NFadeL 157, L157d

'Aux boards
  NFadeLm 101, Light1
  NFadeL 101, Light1a
  NFadeLm 102, Light2
  NFadeL 102, Light2a
  NFadeLm 103, Light3a
  NFadeL 103, Light3
  NFadeLm 104, Light4a
  NFadeL 104, Light4
  NFadeLm 105, Light5a
  NFadeL 105, Light5

  Flash 111, Flasher1
  Flash 112, Flasher2
  Flash 113, Flasher3
  Flash 114, Flasher4
  Flash 115, Flasher5
  Flash 116, Flasher6
  Flash 117, Flasher7
  Flash 118, Flasher8
  Flash 119, Flasher9
  Flash 120, Flasher10

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

'**********************************************************************************************************
'Digital Display
'**********************************************************************************************************
 Dim Digits(40)
 Digits(0)=Array(a00, a05, a0c, a0d, a08, a01, a06, a0f, a02, a03, a04, a07, a0b, a0a, a09, a0e)
 Digits(1)=Array(a10, a15, a1c, a1d, a18, a11, a16, a1f, a12, a13, a14, a17, a1b, a1a, a19, a1e)
 Digits(2)=Array(a20, a25, a2c, a2d, a28, a21, a26, a2f, a22, a23, a24, a27, a2b, a2a, a29, a2e)
 Digits(3)=Array(a30, a35, a3c, a3d, a38, a31, a36, a3f, a32, a33, a34, a37, a3b, a3a, a39, a3e)
 Digits(4)=Array(a40, a45, a4c, a4d, a48, a41, a46, a4f, a42, a43, a44, a47, a4b, a4a, a49, a4e)
 Digits(5)=Array(a50, a55, a5c, a5d, a58, a51, a56, a5f, a52, a53, a54, a57, a5b, a5a, a59, a5e)
 Digits(6)=Array(a60, a65, a6c, a6d, a68, a61, a66, a6f, a62, a63, a64, a67, a6b, a6a, a69, a6e)
 Digits(7)=Array(a70, a75, a7c, a7d, a78, a71, a76, a7f, a72, a73, a74, a77, a7b, a7a, a79, a7e)
 Digits(8)=Array(a80, a85, a8c, a8d, a88, a81, a86, a8f, a82, a83, a84, a87, a8b, a8a, a89, a8e)
 Digits(9)=Array(a90, a95, a9c, a9d, a98, a91, a96, a9f, a92, a93, a94, a97, a9b, a9a, a99, a9e)
 Digits(10)=Array(aa0, aa5, aac, aad, aa8, aa1, aa6, aaf, aa2, aa3, aa4, aa7, aab, aaa, aa9, aae)
 Digits(11)=Array(ab0, ab5, abc, abd, ab8, ab1, ab6, abf, ab2, ab3, ab4, ab7, abb, aba, ab9, abe)
 Digits(12)=Array(ac0, ac5, acc, acd, ac8, ac1, ac6, acf, ac2, ac3, ac4, ac7, acb, aca, ac9, ace)
 Digits(13)=Array(ad0, ad5, adc, add, ad8, ad1, ad6, adf, ad2, ad3, ad4, ad7, adb, ada, ad9, ade)
 Digits(14)=Array(ae0, ae5, aec, aed, ae8, ae1, ae6, aef, ae2, ae3, ae4, ae7, aeb, aea, ae9, aee)
 Digits(15)=Array(af0, af5, afc, afd, af8, af1, af6, aff, af2, af3, af4, af7, afb, afa, af9, afe)
 Digits(16)=Array(b00, b05, b0c, b0d, b08, b01, b06, b0f, b02, b03, b04, b07, b0b, b0a, b09, b0e)
 Digits(17)=Array(b10, b15, b1c, b1d, b18, b11, b16, b1f, b12, b13, b14, b17, b1b, b1a, b19, b1e)
 Digits(18)=Array(b20, b25, b2c, b2d, b28, b21, b26, b2f, b22, b23, b24, b27, b2b, b2a, b29, b2e)
 Digits(19)=Array(b30, b35, b3c, b3d, b38, b31, b36, b3f, b32, b33, b34, b37, b3b, b3a, b39, b3e)

 Digits(20)=Array(b40, b45, b4c, b4d, b48, b41, b46, b4f, b42, b43, b44, b47, b4b, b4a, b49, b4e)
 Digits(21)=Array(b50, b55, b5c, b5d, b58, b51, b56, b5f, b52, b53, b54, b57, b5b, b5a, b59, b5e)
 Digits(22)=Array(b60, b65, b6c, b6d, b68, b61, b66, b6f, b62, b63, b64, b67, b6b, b6a, b69, b6e)
 Digits(23)=Array(b70, b75, b7c, b7d, b78, b71, b76, b7f, b72, b73, b74, b77, b7b, b7a, b79, b7e)
 Digits(24)=Array(b80, b85, b8c, b8d, b88, b81, b86, b8f, b82, b83, b84, b87, b8b, b8a, b89, b8e)
 Digits(25)=Array(b90, b95, b9c, b9d, b98, b91, b96, b9f, b92, b93, b94, b97, b9b, b9a, b99, b9e)
 Digits(26)=Array(ba0, ba5, bac, bad, ba8, ba1, ba6, baf, ba2, ba3, ba4, ba7, bab, baa, ba9, bae)
 Digits(27)=Array(bb0, bb5, bbc, bbd, bb8, bb1, bb6, bbf, bb2, bb3, bb4, bb7, bbb, bba, bb9, bbe)
 Digits(28)=Array(bc0, bc5, bcc, bcd, bc8, bc1, bc6, bcf, bc2, bc3, bc4, bc7, bcb, bca, bc9, bce)
 Digits(29)=Array(bd0, bd5, bdc, bdd, bd8, bd1, bd6, bdf, bd2, bd3, bd4, bd7, bdb, bda, bd9, bde)
 Digits(30)=Array(be0, be5, bec, bed, be8, be1, be6, bef, be2, be3, be4, be7, beb, bea, be9, bee)
 Digits(31)=Array(bf0, bf5, bfc, bfd, bf8, bf1, bf6, bff, bf2, bf3, bf4, bf7, bfb, bfa, bf9, bfe)
 Digits(32)=Array(ac18, ac16, acc1, acd1, ac19, ac17, ac15, acf1, ac11, ac13, ac12, ac14, acb1, aca1, ac10, ace1)
 Digits(33)=Array(ad18, ad16, adc1, add1, ad19, ad17, ad15, adf1, ad11, ad13, ad12, ad14, adb1, ada1, ad10, ade1)
 Digits(34)=Array(ae18, ae16, aec1, aed1, ae19, ae17, ae15, aef1, ae11, ae13, ae12, ae14, aeb1, aea1, ae10, aee1)
 Digits(35)=Array(af18, af16, afc1, afd1, af19, af17, af15, aff1, af11, af13, af12, af14, afb1, afa1, af10, afe1)
 Digits(36)=Array(b9, b7, b0c1, b0d1, b100, b8, b6, b0f1, b2, b4, b3, b5, b0b1, b0a1, b1,b0e1)
 Digits(37)=Array(b109, b107, b1c1, b1d1, b110, b108, b106, b1f1, b102, b104, b103, b105, b1b1, b1a1, b101,b1e1)
 Digits(38)=Array(b119, b117, b2c1, b2d1, b120, b118, b116, b2f1, b112, b114, b113, b115, b2b1, b2a1, b111, b2e1)
 Digits(39)=Array(b129, b127, b3c1, b3d1, b130, b128, b126, b3f1, b122, b3b1, b123, b125, b3b1, b3a1, b121, b3e1)


 Sub DisplayTimer_Timer
    Dim ChgLED, ii, jj, num, chg, stat, obj, b, x
    ChgLED=Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
    If Not IsEmpty(ChgLED)Then
    If DesktopMode = True Then
       For ii=0 To UBound(chgLED)
          num=chgLED(ii, 0) : chg=chgLED(ii, 1) : stat=chgLED(ii, 2)
      if (num < 40) then
              For Each obj In Digits(num)
                   If chg And 1 Then obj.State=stat And 1
                   chg=chg\2 : stat=stat\2
                  Next
      Else
             end if
        Next
     end if
    End If
 End Sub
'**********************************************************************************************************
'**********************************************************************************************************

'Gottlieb Bad Girls
'added by Inkochnito
Sub editDips
  Dim vpmDips : Set vpmDips = New cvpmDips
  With vpmDips
    .AddForm 500,600,"Bad Girls - DIP switches"
    .AddFrame 2,4,190,"Maximum credits",49152,Array("8 credits",0,"10 credits",32768,"15 credits",&H00004000,"20 credits",49152)'dip 15&16
    .AddFrame 2,80,190,"Coin chute 1 and 2 control",&H00002000,Array("seperate",0,"same",&H00002000)'dip 14
    .AddFrame 2,126,190,"Playfield special",&H00200000,Array("replay",0,"extra ball",&H00200000)'dip 22
    .AddFrame 2,172,190,"Game mode",&H10000000,Array("replay",0,"extra ball",&H10000000)'dip 29
    .AddFrame 2,218,600,"Playfield freature control",&H80000000,Array("clear 'Rack' lamps, bumper jackpot starts at 12 hits, 'Rack' special after completing 3 racks",0,"restore 'Rack' lamps, bumper jackpot starts at 8 hits, 'Rack' special after completing 4 racks",&H80000000)'dip 32
    .AddFrame 2,267,600,"Playfield special control",&H40000000,Array("Bozo lights 200K, right upkicker no pool letter, vari-target 10X-50K-special, 'Rerack' special and bonus countdown",0,"Bozo lights 100K, right upkicker spots a pool letter, vari-target 10X-special, 'Rerack' special only",&H40000000)'dip 31
    .AddFrame 205,4,190,"High game to date awards",&H00C00000,Array("not displayed and no award",0,"displayed and no award",&H00800000,"displayed and 2 replays",&H00400000,"displayed and 3 replays",&H00C00000)'dip 23&24
    .AddFrame 205,80,190,"Balls per game",&H01000000,Array("5 balls",0,"3 balls",&H01000000)'dip 25
    .AddFrame 205,126,190,"Replay limit",&H04000000,Array("no limit",0,"one per ball",&H04000000)'dip 27
    .AddFrame 205,172,190,"Novelty",&H08000000,Array("normal",0,"extra ball and replay scores 500K",&H08000000)'dip 28
    .AddFrame 410,4,190,"3rd coin chute credits control",&H20000000,Array("no effect",0,"add 9",&H20000000)'dip 30
    .AddFrame 410,50,190,"High games to date control",&H00000020,Array("no effect",0,"reset high games 2-5 on power off",&H00000020)'dip 6
    .AddFrame 410,96,190,"Auto-percentage control",&H00000080,Array("disabled (normal high score mode)",0,"enabled",&H00000080)'dip 8
    .AddChk 410,147,180,Array("Match feature",&H02000000)'dip 26
    .AddChk 410,162,190,Array("Attract sound",&H00000040)'dip 7
    .AddLabel 200,320,300,20,"After hitting OK, press F3 to reset game with new settings."
    .ViewDips
  End With
End Sub
Set vpmShowDips = GetRef("editDips")


'*********************************************************************
'                 Start of VPX Functions
'*********************************************************************\


'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep

Sub RightSlingShot_Slingshot
  vpmTimer.PulseSw 35
    PlaySoundAtVol SoundFX("right_slingshot",DOFContactors), ActiveBall, 1
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
  vpmTimer.PulseSw 25
    PlaySoundAtVol SoundFX("left_slingshot",DOFContactors), ActiveBall, 1
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
				  ' PlaySound("fx_ballrolling" & b), -1, BallRollVol(BOT(b) )*.8, AudioPan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0, AudioFade(BOT(b) )
        Else ' Ball on raised ramp
          PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) )*.5, AudioPan(BOT(b) ), 0, Pitch(BOT(b) )+50000, 1, 0, AudioFade(BOT(b) )
				  ' PlaySound("fx_ballrolling" & b), -1, BallRollVol(BOT(b) )*.2, AudioPan(BOT(b) ), 0, Pitch(BOT(b) )+50000, 1, 0, AudioFade(BOT(b) )
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
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 2000, AudioPan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub



'*****************************************
' ninuzzu's FLIPPER SHADOWS
'*****************************************

sub FlipperTimer_Timer()
  FlipperLSh.RotZ = LeftFlipper.currentangle
  FlipperRSh.RotZ = RightFlipper.currentangle
  FlipperRSh1.RotZ = RightFlipper1.currentangle

  FlipperT5L.Roty = LeftFlipper.currentangle +240
  FlipperT5R.Roty = RightFlipper.currentangle +120
  FlipperT5R1.Roty = RightFlipper1.currentangle +150
End Sub

'*****************************************
' ninuzzu's BALL SHADOW
'*****************************************
Dim BallShadow
BallShadow = Array (BallShadow1,BallShadow2,BallShadow3,BallShadow4,BallShadow5)

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

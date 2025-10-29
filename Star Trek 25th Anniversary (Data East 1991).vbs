Option Explicit
Randomize

' Thalamus 2018-11-01 : Improved directional sounds
' Thalamus 2018-12-18 : Added FFv2
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


On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Const cGameName="trek_300",UseSolenoids=2,UseLamps=0,UseGI=0,SSolenoidOn="SolOn",SSolenoidOff="SolOff", SCoin="coin"

LoadVPM "01120100", "de.vbs", 3.02

' Thalamus - for Fast Flip v2
NoUpperRightFlipper
NoUpperLeftFlipper


Dim DesktopMode: DesktopMode = Table1.ShowDT
If DesktopMode = True Then 'Show Desktop components
Ramp16.visible=1
Ramp15.visible=1
Primitive13.visible=1
TransReel.visible=1
Else
Ramp16.visible=0
Ramp15.visible=0
Primitive13.visible=0
TransReel.visible=0
End if

'*************************************************************
'Solenoid Call backs
'**********************************************************************************************************
SolCallback(1) = "bsTrough.SolIn"
SolCallback(2) = "bsTrough.SolOut"
SolCallback(3) = "Auto_Plunger"
SolCallback(4) = "dtLBank.SolDropUp"
SolCallback(5) = "SolRampRaise"
SolCallback(6) = "SolRampLower"
SolCallback(7) = "dtRBank.SolDropUp"
SolCallback(8) = "vpmSolSound SoundFX(""Knocker"",DOFKnocker),"
SolCallback(9) = "SetLamp 109,"

SolCallback(11) = "PFGI"
SolCallback(12) = "SetLamp 112,"
'SolCallback(13) = "transporter" ' Backglass
SolCallback(14) = "SubwayEject"
SolCallback(15) = "Solkickback"
SolCallback(16) = "soltarget"

SolCallback(22) = "SolDiv"

SolCallback(25) = "SetLamp 125," '1B
SolCallback(26) = "SetLamp 126," '2B
SolCallback(27) = "SetLamp 127," '3B
SolCallback(28) = "SetLamp 128," '4B
SolCallback(29) = "SetLamp 129," '5B
SolCallback(30) = "SetLamp 130," '6B
SolCallback(31) = "SetLamp 131," '7B
SolCallback(32) = "SetLamp 132," '8B
SolCallback(37) = "SetLamp 137," ' ramp LED strip
SolCallback(38) = "SetLamp 138," ' ramp LED strip
SolCallback(39) = "SetLamp 139," ' ramp LED strip


SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
     If Enabled Then
         PlaySoundAtVol SoundFX("fx_Flipperup",DOFContactors),LeftFlipper,VolFlip:LeftFlipper.RotateToEnd
     Else
         PlaySoundAtVol SoundFX("fx_Flipperdown",DOFContactors),LeftFlipper,VolFlip:LeftFlipper.RotateToStart
     End If
  End Sub

Sub SolRFlipper(Enabled)
     If Enabled Then
         PlaySoundAtVol SoundFX("fx_Flipperup",DOFContactors),RightFlipper,VolFlip:RightFlipper.RotateToEnd
     Else
         PlaySoundAtVol SoundFX("fx_Flipperdown",DOFContactors),RightFlipper,VolFlip:RightFlipper.RotateToStart
     End If
End Sub
'**********************************************************************************************************

'Solenoid Controlled toys
'**********************************************************************************************************

Sub Auto_Plunger(Enabled) 'plunger
  If Enabled Then
    PlungerIM.AutoFire
  End If
End Sub

Sub Solkickback(Enabled) 'kickback
  If Enabled Then
    plunger1.Fire
    playsoundAtVol SoundFX("Popper",DOFContactors), plunger1, 1
  Else
    plunger1.pullback
  End If
End Sub

Sub SolDiv(Enabled)
  If Enabled Then
    s22.Open= 1
    playsound SoundFX("Diverter",DOFContactors)
  Else
    s22.Open= 0
  End If
End Sub

Sub soltarget(Enabled) ' Moving Target
  If Enabled Then
    MovingTargetTimer.enabled=1
  Else
    MovingTargetTimer.enabled=0
  End If
End Sub

Sub SolRampLower(Enabled)
  If Enabled Then
    iRamp.Collidable= 1
    Controller.Switch(39) = 0
    PrimRamp.ObjRotX = 0
        PlaySoundAtVol "fx_Flipperdown", PrimRamp, 1
    iRamp.visible=1
  End If
End Sub

Sub SolRampRaise(Enabled)
  If Enabled Then
    iRamp.Collidable= false
    Controller.Switch(39) = 1
    PrimRamp.ObjRotX = 23
    playsoundAtVol SoundFX("Popper",DOFContactors), PrimRamp, 1
    iRamp.visible=0
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


Sub SubwayEject (Enabled)
  if enabled Then
  PlaySoundAt SoundFX("Solenoid",DOFContactors), sw40
  sw40.kick 0, 35, 1.56
  controller.switch(40) = 0
  end if
End sub

'**********************************************************************************************************

'Initiate Table
'**********************************************************************************************************

Dim bsTrough, bsSaucer, dtLBank, dtRBank, mTransporter

Sub Table1_Init
  vpmInit Me
  On Error Resume Next
    With Controller
    .GameName = cGameName
    If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
    .SplashInfoLine = "Star Trek 25th Anniversary (Data East 1991)"&chr(13)&"You Suck"
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

  PinMAMETimer.Interval = PinMAMEInterval
  PinMAMETimer.Enabled = 1
  vpmNudge.TiltSwitch = 1
  vpmNudge.Sensitivity = 2
  vpmNudge.TiltObj = Array(Bumper1, Bumper2, Bumper3, LeftSlingshot, RightSlingshot)

  Set bsTrough = New cvpmballstack
    bsTrough.InitSw 10,13, 12, 11, 0, 0, 0, 0
    bsTrough.Initkick BallRelease, 80, 6
    bsTrough.InitExitSnd SoundFX("ballrelease",DOFContactors), SoundFX("Solenoid",DOFContactors)
    bsTrough.Balls = 3

  set dtLBank = new cvpmdroptarget
    dtLBank.InitDrop Array(sw20, sw21, sw22, sw23), Array(20, 21, 22, 23)
    dtLBank.Initsnd SoundFX("DTDrop",DOFContactors),SoundFX("DTReset",DOFContactors)

  set dtRBank = new cvpmdroptarget
    dtRBank.InitDrop Array(sw44, sw45, sw46, sw47), Array(44, 45, 46, 47)
    dtRBank.Initsnd SoundFX("DTDrop",DOFContactors),SoundFX("DTReset",DOFContactors)

  set mTransporter = new cvpmMech
  with mTransporter
    .sol1 = 13
    .MType = vpmMechOneSol + vpmMechReverse + vpmMechNONLinear
    .length = 200
    .steps = 9
    .addsw 31, 0, 0
    .addsw 32, 9, 9
    .callback = Getref("UpdateTransporter")
    .start
  end with

  'Moving Target
  movetarget=14
  movedirection=1
  MovingTargetTimer.enabled=false
  PtargetC.objroty=((movetarget*.4)-10)*-1
  PtargetC1.objroty=((movetarget*.4)-10)*-1


  plunger1.pullback
  Controller.Switch(39) = 0
  iRamp.Collidable=1
  divLeft.IsDropped = 1
  divRight.IsDropped = 0

End Sub



'**********************************************************************************************************
'Plunger code
'**********************************************************************************************************

Sub Table1_KeyDown(ByVal KeyCode)

  If keycode = PlungerKey or keycode = LockBarKey Then Controller.Switch(30) = 1
  If KeyDownHandler(keycode) Then Exit Sub
End Sub

Sub Table1_KeyUp(ByVal KeyCode)

  If keycode = PlungerKey or keycode = LockBarKey Then Controller.Switch(30) = 0
  If KeyUpHandler(keycode) Then Exit Sub
End Sub

  ' Impulse Plunger
  dim plungerIM

  Const IMPowerSetting = 40 ' Plunger Power
  Const IMTime = 0.6        ' Time in seconds for Full Plunge
  Set plungerIM = New cvpmImpulseP
  With plungerIM
    .InitImpulseP swPlunger, IMPowerSetting, IMTime
    .Random 0.3
    .Switch 14
    .InitExitSnd SoundFX("Popper",DOFContactors), SoundFX("Solenoid",DOFContactors)
    .CreateEvents "plungerIM"
  End With

'**********************************************************************************************************

 ' Drain hole and kickers
Sub Drain_Hit:bsTrough.addball me : playsoundAtVol"drain" , Drain, 1: End Sub
Sub sw40_Hit: Controller.Switch(40) = 1 : Stopsound "Subway" : playsound "popper_ball" : End Sub

' Droptargets
Sub sw20_Dropped:dtLBank.hit 1:End Sub
Sub sw21_Dropped:dtLBank.hit 2:End Sub
Sub sw22_Dropped:dtLBank.hit 3:End Sub
Sub sw23_Dropped:dtLBank.hit 4:End Sub

Sub sw44_Dropped:dtRBank.hit 1:End Sub
Sub sw45_Dropped:dtRBank.hit 2:End Sub
Sub sw46_Dropped:dtRBank.hit 3:End Sub
Sub sw47_Dropped:dtRBank.hit 4:End Sub

' Rollovers
Sub sw17_Hit:Controller.Switch(17) = 1 : playsoundAtVol"rollover" , ActiveBall, VolRol: End Sub
Sub sw17_UnHit:Controller.Switch(17) = 0:End Sub
Sub sw18_Hit:Controller.Switch(18) = 1 : playsoundAtVol"rollover" , ActiveBall, VolRol: End Sub
Sub sw18_UnHit:Controller.Switch(18) = 0:End Sub
Sub sw50_Hit:Controller.Switch(50) = 1 : playsoundAtVol"rollover" , ActiveBall, VolRol: End Sub
Sub sw50_UnHit:Controller.Switch(50) = 0:End Sub
Sub sw49_Hit:Controller.Switch(49) = 1 : playsoundAtVol"rollover" , ActiveBall, VolRol: End Sub
Sub sw49_UnHit:Controller.Switch(49) = 0:End Sub
Sub sw41_Hit:Controller.Switch(41) = 1 : playsoundAtVol"rollover" , ActiveBall, VolRol: End Sub
Sub sw41_UnHit:Controller.Switch(41) = 0:End Sub
Sub sw42_Hit:Controller.Switch(42) = 1 : playsoundAtVol"rollover" , ActiveBall, VolRol: End Sub
Sub sw42_UnHit:Controller.Switch(42) = 0:End Sub
Sub sw43_Hit:Controller.Switch(43) = 1 : playsoundAtVol"rollover" , ActiveBall, VolRol: End Sub
Sub sw43_UnHit:Controller.Switch(43) = 0:End Sub
Sub sw48_Hit:Controller.Switch(48) = 1 : playsoundAtVol"rollover" , ActiveBall, VolRol: End Sub
Sub sw48_UnHit:Controller.Switch(48) = 0:End Sub

' Stand Up Targets
Sub sw24_Hit:vpmTimer.PulseSw 24:End Sub
Sub sw26_Hit:vpmTimer.PulseSw 26:End Sub

'Subway Triggers
Sub sw27_Hit:Controller.Switch(27) = 1 : playsoundAtVol"Subway" , sw27, 1: End Sub
Sub sw27_UnHit:Controller.Switch(27) = 0:End Sub
Sub sw28_Hit:Controller.Switch(28) = 1 : playsoundAtVol"Subway" , sw28, 1: End Sub
Sub sw28_UnHit:Controller.Switch(28) = 0:End Sub
Sub sw29_Hit:Controller.Switch(29) = 1 : playsoundAtVol"Subway" , sw29, 1: End Sub
Sub sw29_UnHit:Controller.Switch(29) = 0:End Sub
Sub sw52_Hit: Controller.Switch(52) = 1 : End Sub
Sub sw52_UnHit: Controller.Switch(52) = 0 : End Sub

' Ramp Switches
Sub sw36_Hit:vpmTimer.PulseSw 36:End Sub
Sub sw37_Hit:vpmTimer.PulseSw 37:End Sub
Sub sw38_Hit:vpmTimer.PulseSw 38:End Sub

' Bumpers
Sub Bumper1_Hit:vpmTimer.PulseSw 33 : playsoundAtVol SoundFX("fx_bumper1",DOFContactors), ActiveBall, VolBump: End Sub
Sub Bumper2_Hit:vpmTimer.PulseSw 35 : playsoundAtVol SoundFX("fx_bumper1",DOFContactors), ActiveBall, VolBump: End Sub
Sub Bumper3_Hit:vpmTimer.PulseSw 34 : playsoundAtVol SoundFX("fx_bumper1",DOFContactors), ActiveBall, VolBump: End Sub

'Generic Sounds
Sub Trigger1_Hit:PlaySoundAtVol "fx_ballrampdrop", ActiveBall, 1:End Sub
Sub Trigger2_Hit:PlaySoundAtVol "fx_ballrampdrop", ActiveBall, 1:End Sub

' Transporter Reel animation
Sub UpdateTransporter(newpos, speed, lastpos)
  TransReel.Setvalue(newpos)
End Sub

'Ramp Diverter
Sub DivHelp1_Hit : PlaySoundAtVol "rollover" , ActiveBall, VolRol: DiV1.ObjRotz = 30 : End Sub
Sub DivHelp2_Hit :divLeft.IsDropped = 1:divRight.IsDropped = 0 : DiV1.ObjRotz = 0 : End Sub
Sub DivHelp3_Hit : PlaySoundAtVol "rollover" , ActiveBall, VolRol: DiV1.ObjRotz = 30 : End Sub
Sub DivHelp4_Hit :divLeft.IsDropped = 0:divRight.IsDropped = 1 : DiV1.ObjRotz = 60 : End Sub

'***********************
' Swing Target animation
'***********************

Dim MyPi, SwingStep, SwingPos
MyPi = Round(4 * Atn(1), 6) / 90
SwingStep = 0

sub MovingTargetTimer_timer
  if movedirection=1 Then
'   Collection2(movetarget).visible=0
    Collection2(movetarget).collidable=False
    Collection2(movetarget).hashitevent=False
    movetarget=movetarget+1
'   Collection2(movetarget).visible=1
    Collection2(movetarget).collidable=True
    Collection2(movetarget).hashitevent=True

    if movetarget>48 then
      movedirection=0
    end If
    PtargetC.objroty=((movetarget*.4)-10)*-1
    PtargetC1.objroty=((movetarget*.4)-10)*-1
    exit sub
  Else
'   Collection2(movetarget).visible=0
    Collection2(movetarget).collidable=False
    Collection2(movetarget).hashitevent=False

    movetarget=movetarget-1
'   Collection2(movetarget).visible=1
    Collection2(movetarget).collidable=True
    Collection2(movetarget).hashitevent=True

    if movetarget<1 then
      movedirection=1
    end If
  end If
  PtargetC.objroty=((movetarget*.4)-10)*-1
  PtargetC1.objroty=((movetarget*.4)-10)*-1
end sub

Sub Collection2_hit(idx) 'swing target hit
  vpmTimer.PulseSw 25
  PlaySound "target"
  PtargetC.objrotx=1.5
  CenterT1.uservalue=1
  CenterT1.timerenabled=1
end sub

Sub CenterT1_timer
  Select case me.uservalue
    Case 1:
      PtargetC.objrotx=1
    Case 2:
      PtargetC.objrotx=0.5
    Case 3:
      PtargetC.objrotx=0
    Case 4:
      PtargetC.objrotx=-0.5
    Case 5:
      PtargetC.objrotx=0
      me.timerenabled=0
  end Select
  me.uservalue=me.uservalue+1
End sub

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

Sub UpdateLamps
  NFadeL 1, l1
  NFadeL 2, l2
  NFadeL 3, l3
  NFadeL 4, l4
  NFadeL 5, l5
  NFadeL 6, l6
  NFadeL 7, l7
  'NFadeL 8, l8 'Backglass
  NFadeL 9, l9
  NFadeL 10, l10
  NFadeL 11, l11
  NFadeL 12, l12
  NFadeL 13, l13
  NFadeL 14, l14
  'NFadeL 15, l15 'Backglass
  NFadeL 16, l16
  NFadeL 17, l17
  NFadeL 18, l18
  NFadeL 19, l19
  NFadeL 20, l20
  NFadeL 21, l21
  'NFadeL 22, l22 'Backglass
  NFadeL 23, l23
  NFadeL 24, l24
  NFadeL 25, l25
  NFadeL 26, l26
  NFadeObjm 27, P27, "bulbcover1_yellowON", "bulbcover1_yellow"
  NFadeL 27, l27 'Ramp LED
  NFadeL 28, l28
  'NFadeL 29, l29 'Backglass
  'NFadeL 30, l30 'Ball Launch Button
  'NFadeL 31, l31 'Cabinet Start Button
  NFadeL 32, l32
  NFadeL 33, l33
  NFadeL 34, l34
  NFadeL 35, l35
  NFadeObjm 36, P36, "bulbcover1_redON", "bulbcover1_red"
  Flashm 36, f36
  NFadeL 37, l37
  NFadeL 38, l38
  NFadeL 39, l39
  NFadeL 40, l40
  NFadeL 41, l41
  NFadeL 42, l42
  NFadeL 43, l43
  NFadeObjm 44, P44, "bulbcover1_greenON", "bulbcover1_green"
  flash 44, f44 'Backwall LED
  NFadeObjm 45, P45, "bulbcover1_redON", "bulbcover1_red"
  flash 45, f45 'Backwall LED
  NFadeL 46, l46
  NFadeL 47, l47
  NFadeL 48, l48
  NFadeL 49, l49
  NFadeL 50, l50 'Bumpers
  NFadeLm 50, l50a
  NFadeLm 50, l50b
  NFadeLm 50, L50c
  NFadeLm 50, l50d
  NFadeLm 50, L50e
  NFadeL 51, l51
  NFadeL 52, l52

  'crystals
  NFadeObjm 53, P53, "CrystalON", "Crystal"
  Flash 53, F53
  NFadeObjm 54, P54, "CrystalON", "Crystal"
  Flash 54, F54
  NFadeObjm 55, P55, "CrystalON", "Crystal"
  Flash 55, F55
  NFadeObjm 56, P56, "CrystalGreen", "Crystal"
  Flash 56, F56
  NFadeObjm 57, P57, "CrystalON", "Crystal"
  Flash 57, F57
  NFadeObjm 58, P58, "CrystalON", "Crystal"
  Flash 58, F58
  NFadeObjm 59, P59, "CrystalON", "Crystal"
  Flash 59, F59
  NFadeObjm 60, P60, "CrystalON", "Crystal"
  Flash 60, F60
  NFadeObjm 61, P61, "CrystalON", "Crystal"
  Flash 61, F61

  NFadeL 62, l62
  NFadeObjm 63, P63, "bulbcover1_redON", "bulbcover1_red"
  flash 63, f63 'Backwall LED
  NFadeObjm 64, P64, "bulbcover1_greenON", "bulbcover1_green"
  flash 64, f64 'Backwall LED

'Solenoid Controlled Flashers

  NFadeLm 109, f109
  NFadeLm 109, f109a
  NFadeLm 109, f109b
  NFadeLm 109, f109d
  NFadeL 109, f109c

  NFadeLm 112, f112a
  NFadeL 112, f112

  NFadeLm 125, f125a
  NFadeLm 125, f125b
  NFadeL 125, f125

  NFadeObjm 126, P126a, "dome3_clearON", "dome3_clear"   ' Dome
  NFadeObjm 126, P126b, "dome3_redlogoON", "dome3_redlogo"   ' Dome
  NFadeLm 126, f126a
  NFadeL 126, f126b

  NFadeLm 127, f127a
  NFadeL 127, f127

  NFadeLm 128, f128a
  NFadeL 128, f128

  NFadeObjm 129, P129a, "dome3_redON", "dome3_red"   ' Dome
  NFadeObjm 129, P129b, "dome3_redON", "dome3_red"   ' Dome
  NFadeLm 129,f129a
  NFadeL 129, f129b

  NFadeObjm 130, P130a, "dome3_redON", "dome3_red"   ' Dome
  NFadeObjm 130, P130b, "dome3_redON", "dome3_red"   ' Dome
  NFadeLm 130, f130a
  NFadeL 130, f130b

  NFadeL 131, f131

  NFadeL 132, f132

'Tube Lights
  NFadeLm 137, l137a
  NFadeLm 137, l137b
  NFadeLm 137, l137c
  NFadeLm 137, l137d
  NFadeLm 137, l137e
  NFadeLm 137, l137f
  NFadeLm 137, l137g
  NFadeLm 137, l137h
  NFadeLm 137, l137i
  NFadeLm 137, l137j
  NFadeLm 137, l137k
  NFadeL 137, l137l

  NFadeLm 138, l138a
  NFadeLm 138, l138b
  NFadeLm 138, l138c
  NFadeLm 138, l138d
  NFadeLm 138, l138e
  NFadeLm 138, l138f
  NFadeLm 138, l138g
  NFadeLm 138, l138h
  NFadeLm 138, l138i
  NFadeLm 138, l138j
  NFadeLm 138, l138k
  NFadeL 138, l138l

  NFadeLm 139, l139a
  NFadeLm 139, l139b
  NFadeLm 139, l139c
  NFadeLm 139, l139d
  NFadeLm 139, l139e
  NFadeLm 139, l139f
  NFadeLm 139, l139g
  NFadeLm 139, l139h
  NFadeLm 139, l139i
  NFadeLm 139, l139j
  NFadeLm 139, l139k
  NFadeL 139, l139l

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


'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep

Sub RightSlingShot_Slingshot
  vpmTimer.PulseSw 51
    PlaySoundAtVol SoundFX("right_slingshot",DOFContactors), sling1, 1
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.rotx = 20
    RStep = 0
    RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.rotx = 10
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.rotx = 0:RightSlingShot.TimerEnabled = 0:
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
  vpmTimer.PulseSw 19
    PlaySoundAtVol SoundFX("left_slingshot",DOFContactors), sling2, 1
    LSling.Visible = 0
    LSling1.Visible = 1
    sling2.rotx = 20
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.rotx = 10
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.rotx = 0:LeftSlingShot.TimerEnabled = 0:
    End Select
    LStep = LStep + 1
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

  LFLogo.Roty = LeftFlipper.currentangle
  RFLogo.Roty = RightFlipper.currentangle

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
  PlaySound "pinhit_low", 0, Vol(ActiveBall)*VolPi, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Targets_Hit (idx)
  PlaySound "target", 0, Vol(ActiveBall)*VolTarg, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Metals_Thin_Hit (idx)
  PlaySound "metalhit_thin", 0, Vol(ActiveBall)*VolMetal, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Metals_Medium_Hit (idx)
  PlaySound "metalhit_medium", 0, Vol(ActiveBall)*VolMetal, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Metals2_Hit (idx)
  PlaySound "metalhit2", 0, Vol(ActiveBall)*VolMetal, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Gates_Hit (idx)
  PlaySound "gate4", 0, Vol(ActiveBall)*VolGates, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Spinner_Spin
  PlaySoundAtVol "fx_spinner", Spinner, VolSpin
End Sub

Sub Rubbers_Hit(idx)
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 20 then
    PlaySound "fx_rubber2", 0, Vol(ActiveBall)*VolRH, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End if
  If finalspeed >= 6 AND finalspeed <= 20 then
    RandomSoundRubber()
  End If
End Sub

Sub Posts_Hit(idx)
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 16 then
    PlaySound "fx_rubber2", 0, Vol(ActiveBall)*VolPo, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End if
  If finalspeed >= 6 AND finalspeed <= 16 then
    RandomSoundRubber()
  End If
End Sub

Sub RandomSoundRubber()
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySound "rubber_hit_1", 0, Vol(ActiveBall)*VolRH, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 2 : PlaySound "rubber_hit_2", 0, Vol(ActiveBall)*VolRH, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 3 : PlaySound "rubber_hit_3", 0, Vol(ActiveBall)*VolRH, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
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
    Case 1 : PlaySound "flip_hit_1", 0, Vol(ActiveBall)*VolRH, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 2 : PlaySound "flip_hit_2", 0, Vol(ActiveBall)*VolRH, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 3 : PlaySound "flip_hit_3", 0, Vol(ActiveBall)*VolRH, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End Select
End Sub

' Thalamus : Exit in a clean and proper way
Sub Table1_exit()
  On Error Resume Next
  Controller.Pause = False
  Controller.Stop
End Sub


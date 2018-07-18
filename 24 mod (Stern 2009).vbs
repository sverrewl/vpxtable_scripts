Option Explicit
Randomize

' Thalamus 2018-07-18
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Changed from useSolenoids=1 and added InitVpmFFlipsSAM


Const cGameName="twenty4_150",UseSolenoids=2,UseLamps=0,UseGI=0,SSolenoidOn="SolOn",SSolenoidOff="SolOff", SCoin="coin"

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

LoadVPM "01560000", "sam.VBS", 3.10
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
SolCallback(1) = "solTrough"
SolCallback(2) = "solAutofire"
SolCallback(3) = "bsTEject.SolOut" 'safehouse eject
SolCallback(4) = "dtLDrop.SolDropUp" 'Drop Targets
SolCallback(5) = "dtRDrop.SolDropUp" 'Drop Targets
SolCallback(6) = "vpmSolgate GateL,""gate"","
SolCallback(7) = "vpmSolgate GateR,""gate"","
'SolCallback(8) = 'shaker not used
SolCallback(12) = "SolSafeHouse"
SolCallback(13) =	"dtUDrop.SolDropUp" 'Drop Targets
SolCallback(14) = "SolSniper"
SolCallback(22) = "SolPostUp" 'left ramp post
SolCallback(28) = "SolSuitCase" 'suitcase lockup
SolCallback(29) = "sniperpost" 'sniper up post

'Solenoid Controlled Flashers
SolCallback(19) =	"SetLamp 119," 'flash safehouse x3
SolCallback(20) =	"SetLamp 120," 'flash sniper x2
SolCallback(26) =	"SetLamp 126," 'flash pop x3
SolCallback(27) =	"SetLamp 127," 'flash right side under jet?
SolCallback(31) =	"SetLamp 131," 'Orange Dome Slingshots x2
SolCallback(32) =	"SetLamp 132," 'flash left spinner x3

SolCallback(15) = "SolLFlipper"
SolCallback(16) = "SolRFlipper"

Sub SolLFlipper(Enabled)
  If Enabled Then
    PlaySoundAt SoundFXDOF("fx_Flipperup",101,DOFOn,DOFFlippers),LeftFlipper:LeftFlipper.RotateToEnd
  Else
    PlaySoundAt SoundFXDOF("fx_Flipperdown",101,DOFOff,DOFFlippers),LeftFlipper:LeftFlipper.RotateToStart
  End If
End Sub

Sub SolRFlipper(Enabled)
  If Enabled Then
    PlaySoundAt SoundFXDOF("fx_Flipperup",102,DOFOn,DOFFlippers),RightFlipper:RightFlipper.RotateToEnd
  Else
    PlaySoundAt SoundFXDOF("fx_Flipperdown",102,DOFOff,DOFFlippers),RightFlipper:RightFlipper.RotateToStart
  End If
End Sub

'**********************************************************************************************************

'Solenoid Controlled toys
'**********************************************************************************************************

Sub solTrough(Enabled)
  If Enabled Then
    bsTrough.ExitSol_On
    vpmTimer.PulseSw 22
  End If
End Sub

Sub solSniper(Enabled)
  If Enabled Then
    SniperT.Enabled = 1
  End If
End Sub

Sub solSafehouse(Enabled)
  If Enabled Then
    SafeRot = True
  Else
  End If
End Sub

Sub solAutofire(Enabled)
  If Enabled Then
    PlungerIM.AutoFire
    PlaySound SoundFX("Popper",DOFContactors)
  End If
End Sub

Sub SolPostUp(Enabled)
  If Enabled Then
    PostUp.Isdropped= 0
    playsound SoundFX("fx_bumper2",DOFContactors)
    ' UpdateGI 1, 0
  Else
    PostUp.Isdropped= 1
    playsound SoundFX("fx_bumper2",DOFContactors)
    ' UpdateGI 1, 8
  End If
End Sub

'Suitcase post
Sub SolSuitCase(Enabled)
  If Enabled Then
    Post45.IsDropped=True
    SuitRot = True
    scdiv.TransX=40
    playsound SoundFX("fx_bumper2",DOFContactors)
  Else
    Post45.IsDropped=False
    scdiv.TransX=0
    playsound SoundFX("fx_bumper2",DOFContactors)
  End If
End Sub

' Left Ramp post
Sub leftramppost(Enabled)
  If Enabled Then
    RampPost.Isdropped=false
    playsound SoundFX("fx_bumper2",DOFContactors)
  Else
    RampPost.Isdropped=true
    playsound SoundFX("fx_bumper2",DOFContactors)
  End If
End Sub

'Stern-Sega GI
set GICallback = GetRef("UpdateGI")

Sub UpdateGI(no, Enabled)
  If Enabled Then
    DOF 101, DOFOn
    dim xx
    For each xx in GI:xx.State = 1: Next
    PlaySound "fx_relay"
  Else
    DOF 101, DOFOff
    For each xx in GI:xx.State = 0: Next
    PlaySound "fx_relay"
  End If
End Sub

'**********************************************************************************************************

'Initiate Table
'**********************************************************************************************************

Dim bsTrough, bsVUK, visibleLock, bsTEject, dtUDrop, dtLDrop, dtRDrop

Sub Table1_Init
  With Controller
    .GameName = cGameName
    If Err Then MsgBox "Can't start Game " & cGameName & vbNewLine & Err.Description:Exit Sub
      .SplashInfoLine = "24"
      .HandleKeyboard = 0
      .ShowTitle = 0
      .ShowDMDOnly = 1
      .ShowFrame = 0
      .HandleMechanics = 1
      .Hidden = 0
      On Error Resume Next
      .Run GetPlayerHWnd
      If Err Then MsgBox Err.Description
  End With

	PinMAMETimer.Interval = PinMAMEInterval
	PinMAMETimer.Enabled = 1
  vpmNudge.TiltSwitch=-7
  vpmNudge.Sensitivity=1
  vpmNudge.TiltObj=Array(Bumper1,Bumper2,Bumper3,LeftSlingshot,RightSlingshot)

  Set bsTrough = New cvpmBallStack
  bsTrough.InitSw 0, 21, 20, 19, 18, 0, 0, 0
  bsTrough.InitKick BallRelease, 90, 8
  bsTrough.InitExitSnd SoundFX("ballrelease",DOFContactors), SoundFX("Solenoid",DOFContactors)
  bsTrough.Balls = 4

  Set bsTEject = new cvpmBallStack
  bsTEject.InitSaucer sw3, 3, 162, 15
  bsTEject.InitExitSnd SoundFX("Popper",DOFContactors), SoundFX("Solenoid",DOFContactors)

	Set visibleLock = New cvpmVLock
	With visibleLock
		.InitVLock Array(sw43,sw44,sw45),Array(k43, k44, k45), Array(43,44,45)
		.InitSnd SoundFX("Popper",DOFContactors), SoundFX("Solenoid",DOFContactors)
		.ExitDir = 180
		.ExitForce = 0
		.createevents "visibleLock"
	End With

  Set dtUDrop = new cvpmDropTarget
  dtUDrop.Initdrop Array(Sw11,Sw13), Array(11,13)
  dtUDrop.InitSnd SoundFX("DTDrop",DOFDropTargets),SoundFX("DTReset",DOFContactors)

  Set dtLDrop = new cvpmDropTarget
  dtLDrop.Initdrop Array(sw60), Array(60)
  dtLDrop.InitSnd SoundFX("DTDrop",DOFDropTargets),SoundFX("DTReset",DOFContactors)

  Set dtRDrop = new cvpmDropTarget
  dtRDrop.Initdrop Array(sw61), Array(61)
  dtRDrop.InitSnd SoundFX("DTDrop",DOFDropTargets),SoundFX("DTReset",DOFContactors)

	RampPost.Isdropped=true
	RightPost.Isdropped=true
	Post43.IsDropped=true
	Post44.IsDropped=true
	Post45.IsDropped=true
	PostUp.Isdropped = 1

  InitVpmFFlipsSAM
End Sub

'**********************************************************************************************************
'Plunger code
'**********************************************************************************************************

Sub Table1_KeyDown(ByVal Keycode)
  If keycode = PlungerKey Then Plunger.Pullback:PlaySoundAt "plungerpull", Plunger:Controller.Switch(11) = 0
  If vpmKeyDown(keycode) Then Exit Sub
End Sub

Sub Table1_KeyUp(ByVal keycode)
	If vpmKeyUp(keycode) Then Exit Sub
	If Keycode = StartGameKey Then Controller.Switch(16) = 0
	If keycode = PlungerKey Then Plunger.Fire:PlaySoundAt "plunger", Plunger
End Sub

'Auto Plunger
'**********************************************************************************************************

Dim PlungerIM
Const IMPowerSetting = 40
Const IMTime = 0.6
Set plungerIM = New cvpmImpulseP

With plungerIM
  .InitImpulseP swplunger, IMPowerSetting, IMTime
  .Switch 23
  .Random 1.5
  .InitExitSnd "plunger2", "plunger"
  .CreateEvents "plungerIM"
End With

'**********************************************************************************************************

 ' Drain hole and kickers

Sub Drain_Hit:bsTrough.addball me : PlaySoundAt "drain", Drain : End Sub
Sub sw3_Hit:bsTEject.addball 0 : playsound "popper_ball": End Sub

'Drop Targets
Sub Sw11_Dropped:dtUDrop.Hit 1 :End Sub
Sub Sw13_Dropped:dtUDrop.Hit 2 :End Sub
Sub Sw60_Dropped:dtLDrop.Hit 1 :End Sub
Sub Sw61_Dropped:dtRDrop.Hit 1 :End Sub

'Wire Triggers
Sub sw23_Hit:Controller.Switch(23)=1 : PlaySoundAt "rollover", sw23 : End Sub
Sub sw23_unHit:Controller.Switch(23)=0:End Sub
Sub sw24_Hit:Controller.Switch(24)=1 : PlaySoundAt "rollover", sw24 : End Sub
Sub sw24_unHit:Controller.Switch(24)=0:End Sub
Sub sw25_Hit:Controller.Switch(25)=1 : PlaySoundAt "rollover", sw25 : End Sub
Sub sw25_unHit:Controller.Switch(25)=0:End Sub
Sub sw28_Hit:Controller.Switch(28)=1 : PlaySoundAt "rollover", sw28 : End Sub
Sub sw28_unHit:Controller.Switch(28)=0:End Sub
Sub sw29_Hit:Controller.Switch(29)=1 : PlaySoundAt "rollover", sw29 : End Sub
Sub sw29_unHit:Controller.Switch(29)=0:End Sub

Sub sw46_Hit:Controller.Switch(46)=1 : PlaySoundAt "rollover", sw46 : End Sub
Sub sw46_unHit:Controller.Switch(46)=0:End Sub
Sub sw55_Hit:Controller.Switch(55)=1 : PlaySoundAt "rollover", sw55 : End Sub
Sub sw55_unHit:Controller.Switch(55)=0:End Sub
Sub sw56_Hit:Controller.Switch(56)=1 : PlaySoundAt "rollover", sw56 : End Sub
Sub sw56_unHit:Controller.Switch(56)=0:End Sub
Sub sw57_Hit:Controller.Switch(57)=1 : PlaySoundAt "rollover", sw57 : End Sub
Sub sw57_unHit:Controller.Switch(57)=0:End Sub

 'Gate Triggers
Sub sw14_hit:vpmTimer.pulseSw 14 : End Sub
Sub sw54_hit:vpmTimer.pulseSw 54 : End Sub
Sub sw58_hit:vpmTimer.pulseSw 58 : End Sub

'Spinners
Sub sw10_Spin:vpmTimer.PulseSw 10 : PlaySoundAt "fx_spinner", sw10 : End Sub

 'Stand Up Targets
Sub sw1_hit:vpmTimer.pulseSw 1 : End Sub
Sub sw2_hit:vpmTimer.pulseSw 2 : End Sub
Sub sw4_hit:vpmTimer.pulseSw 4 : End Sub
Sub sw5_hit:vpmTimer.pulseSw 5 : End Sub
Sub sw6_hit:vpmTimer.pulseSw 6 : End Sub
Sub sw7_hit:vpmTimer.pulseSw 7 : End Sub
Sub sw8_hit:vpmTimer.pulseSw 8 : End Sub
Sub sw9_hit:vpmTimer.pulseSw 9 : End Sub
Sub sw33_hit:vpmTimer.pulseSw 33 : End Sub
Sub sw34_hit:vpmTimer.pulseSw 34 : End Sub
Sub sw39_hit:vpmTimer.pulseSw 39 : End Sub
Sub sw40_hit:vpmTimer.pulseSw 40 : End Sub
Sub sw41_hit:vpmTimer.pulseSw 41 : End Sub

Sub sw62_Hit
	If sniperstate = true Then
		vpmTimer.PulseSw 62
		SniperT.Enabled = 1
	Else
		vpmTimer.PulseSw 62
	End If
End Sub


'Bumpers
Sub Bumper1_Hit : vpmTimer.PulseSw(30) : PlaySoundAt SoundFXDOF("fx_bumper1",107,DofPulse,DOFContactors), Bumper1: End Sub
Sub Bumper2_Hit : vpmTimer.PulseSw(31) : PlaySoundAt SoundFXDOF("fx_bumper1",108,DofPulse,DOFContactors), Bumper2: End Sub
Sub Bumper3_Hit : vpmTimer.PulseSw(32) : PlaySoundAt SoundFXDOF("fx_bumper1",109,DOfPulse,DOFContactors), Bumper3: End Sub

'SuiteCase Lock buttons
Sub sw43_Hit:Controller.Switch(43)=1:End Sub
Sub sw43_unHit:Controller.Switch(43)=0:End Sub
Sub sw44_Hit:Controller.Switch(44)=1:Post43.IsDropped=False:End Sub 	'suitcase lock mid
Sub sw44_unHit:Controller.Switch(44)=0:Post43.IsDropped=True:End Sub
Sub sw45_Hit:Controller.Switch(45)=1:Post44.IsDropped=False:End Sub 	'suitcase lock top
Sub sw45_unHit:Controller.Switch(45)=0:Post44.IsDropped=True:End Sub

'Generic Sounds
Sub Trigger1_Hit: PlaySoundAt "fx_ballrampdrop", Trigger1 : End Sub
Sub Trigger2_Hit: PlaySoundAt "fx_ballrampdrop", Trigger2 : End Sub
Sub Trigger3_Hit: PlaySoundAt "fx_ballrampdrop", Trigger3 : End Sub

Sub Trigger4_Hit: PlaySoundAt "Wire Ramp", Trigger4 : End Sub

'*************************************************************
'Right Post
'**********************************************************************************************************

sniperpostprim.RotAndTra5=-48

Sub sniperpost(enabled)
  if Enabled then
    RightPost.IsDropped=0
    sniperup.Enabled=1
    'UpdateGI 1, 0
  else
    sniperdown.Enabled=1
    RightPost.IsDropped=1
    'UpdateGI 1, 8
  end if
End Sub

Dim STPos
StPos=0

Sub sniperup_Timer()
  Select Case STPos
    Case 1:sniperpostprim.RotAndTra5=-48
    Case 2:sniperpostprim.RotAndTra5=-38
    Case 3:sniperpostprim.RotAndTra5=-28
    Case 4:sniperpostprim.RotAndTra5=-18
    Case 5:sniperpostprim.RotAndTra5=8
    Case 6:sniperpostprim.RotAndTra5=0:sniperup.Enabled=0
  End Select
 	If STpos<6 then STPos=STpos+1
End Sub

Sub sniperdown_Timer()
  Select Case STPos

    Case 1:sniperpostprim.RotAndTra5=-48:sniperdown.Enabled=0
    Case 2:sniperpostprim.RotAndTra5=-38
    Case 3:sniperpostprim.RotAndTra5=-28
    Case 4:sniperpostprim.RotAndTra5=-18
    Case 5:sniperpostprim.RotAndTra5=8
    Case 6:sniperpostprim.RotAndTra5=0
  End Select
  If STpos>0 Then STPos=STpos-1
End Sub

'*************************************************************
'Sniper
'*************************************************************

Dim sniperstate
sniperstate = False

Sub SniperT_Timer()
  If sniperstate = False then
    If Sniper.rotandtra8 <= 95 then
      Sniper.rotandtra8 = Sniper.rotandtra8 + 1
    Else
      SniperT.Enabled = False
      sniperstate = True
    End If
  Else
    If Sniper.rotandtra8 => 25 then
      Sniper.rotandtra8 = Sniper.rotandtra8 - 1
    Else
      SniperT.Enabled = False
      sniperstate = False
    End If
  End If
End Sub

'*************************************************************
'SafeHouse
'*************************************************************

Dim SafeRot:SafeRot = False

Sub SafehouseT_Timer()
  If SafeRot = True and Safehouse.RotY >= 180 then
    DOF 201, DOFOn
    Safehouse.RotY = Safehouse.RotY - 3
    DOF 201, DOFOff
  End If
  If SafeRot = False and Safehouse.RotY <= 270 then
    DOF 201, DOFOn
    Safehouse.RotY = Safehouse.RotY + 3
    DOF 201, DOFOff
  End If
  If Safehouse.RotY = 180 then SafeRot = False
End Sub

'*************************************************************
'SuitCase
'*************************************************************

Dim SuitRot:SuitRot = False

Sub SuitcaseT_Timer()
  If SuitRot = True and Suitcase.RotAndTra7 >= 320 then
    DOF 201, DOFOn
    Suitcase.RotAndTra7 = Suitcase.RotAndTra7 - 1
    DOF 201, DOFOff
  End If
  If SuitRot = False and Suitcase.RotAndTra7 <= 358 then
    DOF 201, DOFOn
    Suitcase.RotAndTra7 = Suitcase.RotAndTra7 + 1
    DOF 201, DOFOff
  End If
  If Suitcase.RotAndTra7 = 320 then SuitRot = False
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
LampTimer.Interval = 5  'lamp fading speed
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
  NFadeL 3, l3
  NFadeL 4, l4
  NFadeL 5, l5
  NFadeL 6, l6
  NFadeL 7, l7
  NFadeL 8, l8
  NFadeL 10, l10
  NFadeL 11, l11
  NFadeL 12, l12
  NFadeL 13, l13
  NFadeL 17, l17
  NFadeL 18, l18
  NFadeL 19, l19
  NFadeL 20, l20
  NFadeL 21, l21
  NFadeL 22, l22
  NFadeL 23, l23
  NFadeL 24, l24
  NFadeL 25, l25
  NFadeL 26, l26
  NFadeL 27, l27
  NFadeL 28, l28
  NFadeL 29, l29
  NFadeL 30, l30
  NFadeL 31, l31
  NFadeL 32, l32
  NFadeL 33, l33
  NFadeL 34, l34
  NFadeL 35, l35
  NFadeL 36, l36
  NFadeL 37, l37
  NFadeL 38, l38
  NFadeL 39, l39
  NFadeL 40, l40
  NFadeL 41, l41
  NFadeL 42, l42
  NFadeL 43, l43
  NFadeL 44, l44
  NFadeL 45, l45
  NFadeL 46, l46
  NFadeL 47, l47
  NFadeL 48, l48
  NFadeL 49, l49
  NFadeL 50, l50
  NFadeL 51, l51
  NFadeL 52, l52
  NFadeL 53, l53
  NFadeL 54, l54
  NFadeL 55, l55
  NFadeL 56, l56
  NFadeL 57, l57
  NFadeL 59, l59
  NFadeObjm 60, P60, "lampbulbON", "lampbulb"
  NFadeL 60, l60 'top left bumper
  NFadeObjm 61, P61, "lampbulbON", "lampbulb"
  NFadeL 61, l61 'right bumper
  NFadeObjm 62, P62, "lampbulbON", "lampbulb"
  NFadeL 63, l63
  NFadeL 64, l64
  NFadeL 65, l65
  NFadeL 66, l66
  NFadeL 67, l67
  NFadeL 68, l68
  NFadeL 69, l69
  NFadeL 70, l70
  NFadeL 71, l71
  NFadeL 72, l72
  NFadeL 73, l73
  NFadeL 74, l74
  NFadeL 75, l75
  NFadeL 76, l76
  NFadeL 77, l77
  NFadeL 78, l78
  NFadeL 79, l79
  NFadeL 80, l80

  'Solenoid Controlled Flashers
  NFadeObjm 119, Safehouse, "safehousetexture_ON", "safehousetexture"
  NFadeObjm 119, Primitive11, "safehousetexture_ON", "safehousetexture"
  NFadeLm 119, f119a 'safehouse
  NFadeL 119, f119 'safehouse

  NFadeObj 120, Sniper, "SniperTexture_ON", "SniperTexture"

  NFadeL 126, f126 'PF light

  NFadeObjm 127, l127, "flasher_red_on", "flasher_red"
  NFadeL 127, f127 'right side red dome

  NFadeObjm 131, l131, "flasher_orange_on", "flasher_orange"
  NFadeObjm 131, l131a, "flasher_orange_on", "flasher_orange"
  NFadeLm 131, f131 'orange flasher left sling
  NFadeL 131, f131a 'orange flasher left sling

  NFadeObjm 132, l132, "flasher_red_on", "flasher_red"
  NFadeL 132, f132 'left red dome
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

'************************************************************************
'	Start of VPX functions
'************************************************************************

'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep

Sub RightSlingShot_Slingshot
  vpmTimer.PulseSw 27
  PlaySound SoundFX("right_slingshot",DOFContactors), 0,1, 0.05,0.05 '0,1, AudioPan(RightSlingShot), 0.05,0,0,1,AudioFade(RightSlingShot)
  RSling.Visible = 0
  RSling1.Visible = 1
  sling1.TransZ = -20
  RStep = 0
  RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
  Select Case RStep
    Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.TransZ = -10
    Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.TransZ = 0:RightSlingShot.TimerEnabled = 0:
  End Select
  RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
  vpmTimer.PulseSw 26
  PlaySound SoundFX("left_slingshot",DOFContactors), 0,1, -0.05,0.05 '0,1, AudioPan(LeftSlingShot), 0.05,0,0,1,AudioFade(LeftSlingShot)
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

Sub PlaySoundAtVol(sound, tableobj, Vol)
  PlaySound sound, 1, Vol, Pan(tableobj), 0,0,0, 1, AudioFade(tableobj)
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

Function AudioFade(ball) ' Can this be together with the above function ?
  Dim tmp
  tmp = ball.y * 2 / Table1.height-1
  If tmp > 0 Then
    AudioFade = Csng(tmp ^10)
  Else
    AudioFade = Csng(-((- tmp) ^10) )
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
  If Table1.VersionMinor > 3 OR Table1.VersionMajor > 10 Then
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 200, Pan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
  Else
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 200, Pan(ball1), 0, Pitch(ball1), 0, 0
  End if
End Sub

'*****************************************
'	ninuzzu's	FLIPPER SHADOWS
'*****************************************

sub FlipperTimer_Timer()
  FlipperLSh.RotZ = LeftFlipper.currentangle
  FlipperRSh.RotZ = RightFlipper.currentangle

  sw14Prim.Rotz = sw14.Currentangle
  sw54Prim.Rotz = sw54.Currentangle
  sw58Prim.Rotz = sw58.Currentangle
End Sub

'*****************************************
'	ninuzzu's	BALL SHADOW
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
        BallShadow(b).X = ((BOT(b).X) - (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/7)) + 6
      Else
        BallShadow(b).X = ((BOT(b).X) + (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/7)) - 6
      End If
      ballShadow(b).Y = BOT(b).Y + 12
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

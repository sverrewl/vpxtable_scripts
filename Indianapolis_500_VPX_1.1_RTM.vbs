' Indianapolis 500 - Bally 1995
 ' JPSalas - VP9 v1.0
 ' Turbo Handler Code by Dorsola

'VP10 Conversion / Enchancement Dozer - August 2017.

' Thalamus 2018-07-24
' Table has its own "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Changed UseSolenoids=1 to 2
' No special SSF tweaks yet.

 Option Explicit
 Randomize
 Const UseVPMModSol = 1

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

'/////////// TABLE OPTIONS //////////////


Const ToyMod = 1

'Render a rotating Indy Car sign up the back left of the Playfield.

Const Rotating_Track = 1
'Rotate the track under the upper Indy Car Toy.

Const Track_Type = 3
' 1=Original White / 2=Asphalt with line / 3=Asphalt with broken line.

Const Sloppy_Turbo_Vuk = 0
' With this on, the VUK hole to feed the turbo will catch balls less precisely. The real
' game suffers from random ball fails when shot into this VUK. Turn off for precise kicker grabs.

Const Turbo_Reflection = 1
'Render a relection of the turbo on the left sidewall.

Const Flipper_Bat_Color = 2

' 1 = Blue / 2 = Black

Const GridLamp = 1
'Change the color of the lamps in the grid light targets on the playfield.

' 1=Red / 2=Green / 3=Blue / 4 = Yellow / 5=Purple / 6=Cyan / 7=Orange

Const TurboLamp = 1 ' 1=Red / 2=Green / 3=Blue / 4 = Yellow / 5=Purple / 6=Cyan / 7=Orange
'Change the color of the lamp underneath the turbo impellor on the turbo mech.

' 1=Red / 2=Green / 3=Blue / 4 = Yellow / 5=Purple / 6=Cyan / 7=Orange

'/////////////////////////////////////////////////////////////////////

Const Dozer_Cab = 0
' Leave this OFF!

'//////////// END OF OPTIONS ////////////////////////////////////////


Dim Ballsize,BallMass
BallSize = 50
BallMass = (Ballsize^3)/125000

Dim VarHidden, UseVPMDMD
If table1.ShowDT = true then
UseVPMDMD = true
VarHidden = 1
'Ramp19.visible = 1
'Ramp18.visible = 1
else
'Ramp19.visible = 1
'Ramp18.visible = 1
UseVPMDMD = false
VarHidden = 0
end if

If toymod = 1 Then
modtoy.enabled = 1
track1.visible = 1
indy_shaft1.visible = 1
modref.visible = 1
TopCar_Base2.visible = 1
Prim_Spot3.visible = 1
lamp_post.visible = 1
End If

If Flipper_Bat_Color = 1 Then
batleft.image = "flipper_white_blue"
batright.image = "flipper_white_blue"
batright1.image = "flipper_white_blue"
else
batleft.image = "flipper_white_black"
batright.image = "flipper_white_black"
batright1.image = "flipper_white_black"
End If

If Track_Type = 1 Then
track.image = "track_white"
End If

If Track_Type = 2 Then
track.image = "track_line"
End If

If Track_Type = 3 Then
track.image = "track_broken"
End If

If Turbo_Reflection = 1 Then
TREF.visible = 1
Turbo_Side_Flash2.opacity = 20
else
TREF.visible = 0
Turbo_Side_Flash2.opacity = 0
End If

If Sloppy_Turbo_VUK = 1 Then
SW62.enabled = 0
else
SW62.enabled = 1
End If

Dim Bulb

If gridlamp = 1 Then
for each bulb in tlights
bulb.Color=RGB(255,0,0)
next
End If

If gridlamp = 2 Then
for each bulb in tlights
bulb.Color=RGB(0,255,0)
next
End If

If gridlamp = 3 Then
for each bulb in tlights
bulb.Color=RGB(0,0,255)
next
End If

If gridlamp = 4 Then
for each bulb in tlights
bulb.Color=RGB(255,255,0)
next
End If

If gridlamp = 5 Then
for each bulb in tlights
bulb.Color=RGB(255,0,255)
next
End If

If gridlamp = 6 Then
for each bulb in tlights
bulb.Color=RGB(0,255,255)
next
End If

If gridlamp = 7 Then
for each bulb in tlights
bulb.Color=RGB(255,102,0)
next
End If

If turbolamp = 1 Then
for each bulb in tblights
bulb.Color=RGB(255,0,0)
next
End If

If turbolamp = 2 Then
for each bulb in tblights
bulb.Color=RGB(0,255,0)
next
End If

If turbolamp = 3 Then
for each bulb in tblights
bulb.Color=RGB(0,0,255)
next
End If

If turbolamp = 4 Then
for each bulb in tblights
bulb.Color=RGB(255,255,0)
next
End If

If turbolamp = 5 Then
for each bulb in tblights
bulb.Color=RGB(255,0,255)
next
End If

If turbolamp = 6 Then
for each bulb in tblights
bulb.Color=RGB(0,255,255)
next
End If

If turbolamp = 7 Then
for each bulb in tblights
bulb.Color=RGB(255,102,0)
next
End If

 'LoadVPM "01560000", "wpc.vbs", 3.26
LoadVPM "02800000", "WPC.VBS", 3.55

 ' Wob: Needed for Fast Flips
 NoUpperLeftFlipper

 ' Init Table
 Const cGameName = "i500_11r"
 Const UseSolenoids = 2
 Const UseLamps = 1
 Const UseGI = 0
 Const UseSync = 0
 Const FlasherTest = 0


 Dim x, bsTrough, bsLE, bsUE, bsTBP, bsBBP, bump1, bump2, bump3, BallInShooterLane, TopCarPos, MSpinMagnet, MSpinMagnet1

 ' Standard Sounds
 Const SSolenoidOn = "Solenoid"    'Solenoid activates
 Const SSolenoidOff = ""           'Solenoid deactivates
 Const SFlipperOn = "FlipperUp"    'Flipper activated
 Const SFlipperOff = "FlipperDown" 'Flipper deactivated
 Const SCoin = "Coin"              'Coin inserted

 '************
 ' Table init.
 '************

 Sub Table1_Init
     Dim i
     With Controller
          vpmInit Me
         .GameName = cGameName
         If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description:Exit Sub
         .SplashInfoLine = "Indianapolis 500 - (Bally 1995)"
         .HandleKeyboard = 0
         .ShowTitle = 0
         .ShowDMDOnly = 1
         .ShowFrame = 0
         .HandleMechanics = 1
         .Hidden = VarHidden
         .Games(cGameName).Settings.Value("sound") = 1
         '.Games(cGameName).Settings.Value("rol") = 0 'Do Not Rotate
         On Error Resume Next
         .Run GetPlayerHWnd
         If Err Then MsgBox Err.Description
         On Error Goto 0
     End With

     Controller.Switch(22) = 1 'door closed
     Controller.Switch(24) = 0 'door always closed

     ' Main Timer init
     PinMAMETimer.Interval = PinMAMEInterval
     PinMAMETimer.Enabled = 1
     'StopShake 'StartShake

     ' Nudging
     vpmNudge.TiltSwitch = 14
     vpmNudge.Sensitivity =5
     vpmNudge.TiltObj = Array(Bumper1, Bumper2, Bumper3, LeftSlingshot, RightSlingshot)

    Set bsTrough = New cvpmTrough
    With bsTrough
		.Size = 4
		.InitSwitches Array(42, 43, 44, 45)
		.InitExit BallRelease, 60, 6
		'.InitEntrySounds "fx_drain", SoundFX(SSolenoidOn,DOFContactors), SoundFX(SSolenoidOn,DOFContactors)
		'.InitExitSounds  SoundFX(SSolenoidOn,DOFContactors), SoundFX("BallRelease",DOFContactors)
		.Balls = 4
		'.CreateEvents "bsTrough", Drain
    End With

    Set mSpinMagnet = New cvpmMagnet
    With mSpinMagnet
        .InitMagnet SpinMagnet, 70
        '.Solenoid = 35 'own solenoid sub
        .GrabCenter = 0
        .Size = 100
        .CreateEvents "mSpinMagnet"
    End With

     ' Lower Eject
     Set bsLE = New cvpmBallStack
     'bsLE.InitSaucer sw65, 65, 333, 25
     bsLE.InitSaucer sw65, 65, 328+(RND*5), 23+(RND*2)
     'bsLE.InitExitSnd "popper", "Solenoid"

     ' Upper Eject
     Set bsUE = New cvpmBallStack
     bsUE.InitSaucer sw64, 64, 50, 19
     'bsUE.InitExitSnd "popper", "Solenoid"
     bsUE.KickAngleVar = 4
     bsUE.KickForceVar = 4

     ' Top Ball Popper
     Set bsTBP = New cvpmBallStack
     bsTBP.InitSaucer sw61, 61, 0, 0
     bsTBP.InitKick sw61a, 140, 9
     'bsTBP.InitExitSnd "Popper", "Popper"

     ' Bottom Ball Popper
     Set bsBBP = New cvpmBallStack
     bsBBP.InitSaucer sw62, 62, 0, 0
     bsBBP.InitKick sw62a, 330, 16
     'bsBBP.InitExitSnd "Popper", "Popper"

       DiverterON.IsDropped = 1':CarOn.alpha = 0
     BallInShooterLane = 0
     TopCarPos = 1
     Plunger1.Pullback
     Init_Turbo
 End Sub

 Sub Table1_Paused:Controller.Pause = True:End Sub
 Sub Table1_unPaused:Controller.Pause = False:End Sub

 '****
 'Keys
 '****

 Sub Table1_KeyDown(ByVal Keycode)
    If keycode = LeftTiltKey Then Nudge 90, 5
    If keycode = RightTiltKey Then Nudge 270, 5
    If keycode = CenterTiltKey Then Nudge 0, 6
     If vpmKeyDown(keycode) Then Exit Sub
     If keycode = StartGameKey Then Controller.Switch(13) = 1
     If keycode = keyFront Then Controller.Switch(23) = 1
     If keycode = PlungerKey Then
         Controller.Switch(11) = 1
         If BallInShooterLane = 1 Then
             'Playsound "plunger2",0,1,1.0,0
              PlaySoundAt "plunger2",sw25
         Else
             'PlaySound "plunger",0,1,1.0,0
             'PlaySoundAt "plunger",sw25
         End If
	 End If
     If keycode = KeyRules Then Rules
 End Sub

 Sub Table1_KeyUp(ByVal Keycode)
     If vpmKeyUp(keycode) Then Exit Sub
     If keycode = StartGameKey Then Controller.Switch(13) = 0
     If keycode = keyFront Then Controller.Switch(23) = 0
     If keycode = PlungerKey Then Controller.Switch(11) = 0
 End Sub




 '*********
 'Solenoids
 '*********
 SolCallback(1) = "SolPlungBall"
 SolCallback(2) = "SolTopPopper"
 SolCallback(3) = "UpMidSaucer"
 SolCallback(4) = "slot_kicker"
 SolCallback(5) = "SolBottomPopper"
 SolCallback(7) = "Knocker"
 SolCallback(13) = "SolBallRelease"
 SolModCallback(14) = "Sol14"
 SolModCallback(15) = "Sol15"
 SolModCallback(16) = "Sol16"
 SolCallback(18) = "SolTopCar"
 SolModCallback(19) = "Sol19"
 SolModCallback(20) = "Sol20"
 SolModCallback(21) = "Sol21"
 SolModCallback(22) = "Sol22"
 SolModCallback(23) = "Sol23"
 SolModCallback(24) = "Sol24"
 SolModCallback(25) = "Sol25"
 SolModCallback(26) = "Sol26"
 SolModCallback(27) = "Sol27"
 SolModCallback(28) = "Sol28"
 SolCallback(36) = "SolDiverterHold"

 ' Solenoid Subs

Sub Knocker(Enabled)
If enabled Then
PlaySoundAt SoundFX("Knocker",DOFKnocker), sw61a
End If
End Sub

Sub Sol19(Enabled)
F19a.Intensity = Enabled / 10
If enabled > 1 Then
F19.State = 1
else
F19.State = 0
End If
End Sub

Sub Sol20(Enabled)
F20a.Intensity = Enabled / 10
If enabled > 1 Then
F20.State = 1
else
F20.State = 0
End If
End Sub

Sub Sol21(Enabled)
F21a.Intensity = Enabled / 10
If enabled > 1 Then
F21.State = 1
F21r.State = 1
else
F21.State = 0
F21r.State = 0
End If
End Sub

Sub Sol22(Enabled)
F22a.Intensity = Enabled / 10
If enabled > 1 Then
F22.State = 1
else
F22.State = 0
End If
End Sub

Sub Sol26(Enabled)
F26_Lamp1.Intensity = Enabled / 10
F26_Lamp2.Intensity = Enabled / 20
F26_Side_Flash1.opacity = Enabled / 2
End Sub

'/////////////

Sub Sol16(Enabled)
F16_Lamp1.Intensity = Enabled / 10
F16_Lamp2.Intensity = Enabled / 20
F16_Side_Flash1.opacity = Enabled / 2
End Sub

Sub Sol14(Enabled)
F14_Lamp1.Intensity = Enabled / 10
F14_Lamp2.Intensity = Enabled / 10
F14_Lamp3.Intensity = Enabled / 10

If toymod = 1 Then
t1.Intensity = Enabled / 10
Turbo_Side_Flash3.opacity = Enabled / 2
Turbo_Side_Flash1.opacity = Enabled / 2
End If

If Enabled > 1 AND toymod = 1 Then
Track1.image = "disc2Db_asp2log"
Track3.visible = 1
Indy_Shaft3.visible = 1
else
Track1.image = "disc2Db_asp2log_dark"
Track3.visible = 0
Indy_Shaft3.visible = 0
End If

F14_Side_Flash1.opacity = Enabled / 2
If enabled > 1 Then
F14_Side_Flash2.visible = 1
else
F14_Side_Flash2.visible = 0
End If
End Sub

Sub Sol15(Enabled)
F15_Lamp1.Intensity = Enabled / 10
F15_Lamp2.Intensity = Enabled / 10
F15_Lamp3.Intensity = Enabled / 10
F15_Side_Flash1.opacity = Enabled / 2
F15_Side_Flash2.opacity = Enabled / 2

If toymod = 1 Then
t1.Intensity = Enabled / 10
Turbo_Side_Flash3.opacity = Enabled / 2
Turbo_Side_Flash1.opacity = Enabled / 2
End If

If enabled > 1 Then
F15_Side_Flash3.visible = 1
else
F15_Side_Flash3.visible = 0
End If

If Enabled > 1 AND toymod = 1 Then
Track3.visible = 1
Indy_Shaft3.visible = 1
Track1.image = "disc2Db_asp2log"
else
Track3.visible = 0
Indy_Shaft3.visible = 0
Track1.image = "disc2Db_asp2log_dark"
End If
End Sub

Sub Sol23(Enabled)
Flash23_Lamp1.Intensity = Enabled / 10
Flash23_Lamp2.Intensity = Enabled / 10
Flash23_Lamp3.Intensity = Enabled / 10
Flash23_Lamp4.Intensity = Enabled / 10
End Sub

Sub Sol24(Enabled)
Flash24_Lamp1.Intensity = Enabled / 10
Flash24_Lamp2.Intensity = Enabled / 10
Flash24_Lamp3.Intensity = Enabled / 10
Flash24_Lamp4.Intensity = Enabled / 10
End Sub

Sub Sol25(Enabled)
Flash25_Lamp1.Intensity = Enabled / 10
Flash25_Lamp2.Intensity = Enabled / 10
Flash25_Lamp3.Intensity = Enabled / 10
Flash25_Lamp4.Intensity = Enabled / 10
End Sub

Sub Sol27(Enabled)
F27_Lamp1.Intensity = Enabled / 10
F27_Lamp2.Intensity = Enabled / 10
F27_Lamp3.Intensity = Enabled / 10
F27_Lamp4.Intensity = Enabled / 10
F27_Lamp5.Intensity = Enabled / 20
F27_Lamp6.Intensity = Enabled / 10
F27_Lamp7.Intensity = Enabled / 10
F27_Lamp8.Intensity = Enabled / 10
F27_Lamp9.Intensity = Enabled / 10
F27_Lamp10.Intensity = Enabled / 10
F27_Lamp11.Intensity = Enabled / 20
F27_LowBoy.Intensity = Enabled / 20
F27_Flash.opacity = Enabled / 2
F27_Side_Flash.opacity = Enabled / 2
F27_Side_Flash1.opacity = Enabled / 2
End Sub

Sub Sol28(Enabled)
F28_LowBoy.Intensity = Enabled / 20
F28_Lamp1.Intensity = Enabled / 10
F28_Lamp2.Intensity = Enabled / 10
F28_Lamp3.Intensity = Enabled / 10
F28_Lamp4.Intensity = Enabled / 10
F28_Lamp5.Intensity = Enabled / 20
F28_Flash1.opacity = Enabled / 2
F28_Side_Flash.opacity = Enabled / 2

If enabled > 1 Then
F17s4.visible = 1
else
F17s4.visible = 0
End If
End Sub

 Sub UpMidSaucer(Enabled)
 If enabled Then
 bsUE.ExitSol_on
 PlaySoundAt SoundFX("solenoid",DOFContactors),sw64
 End If
 End Sub

 Sub SolPlungBall(Enabled)
     If Enabled Then
         'sw25.kick 0, 35
         Plunger1.Fire
         PlaySoundat SoundFX("popper",DOFContactors),sw25
         PlaySoundat "rail_low_slower", sw25
         'Plunger1.PullBack
     End If
 End Sub

 Sub SolBallRelease(Enabled)
     If Enabled Then
        bsTrough.ExitSol_On
         PlaySoundat SoundFX("ballrel",DOFContactors),BallRelease
         If bsTrough.Balls Then
         vpmTimer.PulseSw 41
         End If
     End If
 End Sub

 Sub SolDiverterHold(Enabled)
     If Enabled Then
         PlaySoundat SoundFX("Solenoid",DOFContactors),diverterhold
         ibpos = 1:Indy_Lower.enabled = 1
         DiverterOff.IsDropped = 1
         'carr.state=ABS(carr.state-1)
         DiverterOn.IsDropped = 0
         If Dozer_Cab = 1 Then
         DOF 166,2
         End If
     Else
         PlaySoundat SoundFX("Solenoid",DOFContactors),diverterhold
         ibpos = 2:Indy_Lower.enabled = 1
         DiverterOff.IsDropped = 0
         'carr.state=ABS(carr.state-1)
         DiverterOn.IsDropped = 1
         DiverterHold.Kick 180, 3
         If Dozer_Cab = 1 Then
         DOF 166,2
         End If
     End If
 End Sub


'Lower Racecar Diverter Code.

Dim ibpos

ibpos = 1

Sub Indy_Lower_Timer()
Select Case ibpos
Case 1:
If Indy_Bottom.objrotz <= -20 Then
Indy_Bottom.objrotz = -20
TL1.State = 2
TL2.State = 2
me.enabled = 0
End If
Indy_Bottom.objrotz = Indy_Bottom.objrotz - 1
Case 2:
If Indy_Bottom.objrotz => 0 Then
Indy_Bottom.objrotz = 0
TL1.State = 0
TL2.State = 0
me.enabled = 0
End If
Indy_Bottom.objrotz = Indy_Bottom.objrotz + 1
End Select
End Sub

'Upper Right Slot Kicker Out

Sub slot_kicker(enabled)
If enabled Then
Ukick = 1
upper_kicker.enabled = 1
bsle.exitSol_On
PlaySoundat SoundFX("Solenoid",DOFContactors),sw65
End If
End Sub

Dim ukick

Sub Upper_Kicker_Timer()
Select Case ukick
Case 1:
If Hammer.RotZ => 45 Then
ukick = 2
End If
Hammer.RotZ = Hammer.RotZ + 1
Case 2:
If Hammer.RotZ <= 0 Then
Hammer.RotZ = 0
me.enabled = 0
End If
Hammer.RotZ = Hammer.RotZ - 1
End Select
End Sub

'******************
' RealTime Updates
'******************

Set MotorCallback = GetRef("GameTimer")

Sub GameTimer
BallShadowUpdate
batleft.objrotz = LeftFlipper.CurrentAngle + 1
batright.objrotz = RightFlipper.CurrentAngle + 1
batright1.objrotz = RightFlipper2.CurrentAngle + 1
batleftshadow.objrotz = batleft.objrotz
batrightshadow.objrotz  = batright.objrotz
batrightshadow1.objrotz  = RightFlipper2.CurrentAngle + 1
End Sub

Sub modtoy_timer()
Track1.roty = Track1.roty + 1
'Track2.roty = Track1.roty + 1
Track3.roty = Track1.roty + 1
modref.roty = modref.roty + 1
'modref1.roty = modref1.roty + 1
Indy_shaft1.objrotz = Indy_shaft1.objrotz - 1
TopCar_Base2.objrotz = TopCar_Base2.objrotz - 1
End Sub

 '**************
 ' Flipper Subs
 '**************

 SolCallback(sLRFlipper) = "SolRFlipper"
 SolCallback(sLLFlipper) = "SolLFlipper"


'******************************************
' Use FlipperTimers to call div subs
'******************************************


Sub SolLFlipper(Enabled)
     If Enabled Then
         PlaySoundatvol SoundFX("left_flipper_up", DOFFlippers),leftflipper, 0.5
		 LeftFlipper.RotateToEnd
     Else
		 PlaySoundatvol SoundFX("left_flipper_down", DOFFlippers),leftflipper, 0.5
         LeftFlipper.RotateToStart
     End If
 End Sub


Sub SolRFlipper(Enabled)
     If Enabled Then
         PlaySoundatvol SoundFX("right_flipper_up", DOFFlippers),rightflipper, 0.5
         PlaySoundatvol SoundFX("right_flipper_up", DOFFlippers),rightflipper2, 0.1
         RightFlipper.RotateToEnd:RightFlipper2.RotateToEnd
     Else
		 PlaySoundatvol SoundFX("right_flipper_down", DOFFlippers),rightflipper, 0.5
         PlaySoundatvol SoundFX("right_flipper_down", DOFFlippers),rightflipper2, 0.1
		 RightFlipper.RotateToStart:RightFlipper2.RotateToStart
     End If
 End Sub

 ' Lanes
 Sub sw15_Hit:Playsoundatball "sensor":Controller.Switch(15) = 1:End Sub
 Sub sw15_UnHit:Controller.Switch(15) = 0:End Sub

 Sub sw16_Hit:Playsoundatball "sensor":Controller.Switch(16) = 1:End Sub
 Sub sw16_UnHit:Controller.Switch(16) = 0:End Sub

 Sub sw17_Hit:Playsoundatball "sensor":Controller.Switch(17) = 1:End Sub
 Sub sw17_UnHit:Controller.Switch(17) = 0:End Sub

 Sub sw18_Hit:Playsoundatball "sensor":Controller.Switch(18) = 1:End Sub
 Sub sw18_UnHit:Controller.Switch(18) = 0:End Sub

 Sub sw37_Hit:Playsoundatball "sensor":Controller.Switch(37) = 1:End Sub
 Sub sw37_UnHit:Controller.Switch(37) = 0:End Sub

 Sub sw38_Hit:Playsoundatball "sensor":Controller.Switch(38) = 1:End Sub
 Sub sw38_UnHit:Controller.Switch(38) = 0:End Sub

 Sub sw51_Hit:Playsoundatball "sensor":Controller.Switch(51) = 1:End Sub
 Sub sw51_UnHit:Controller.Switch(51) = 0:End Sub

 Sub sw52_Hit:Playsoundatball "sensor":Controller.Switch(52) = 1:End Sub
 Sub sw52_UnHit:Controller.Switch(52) = 0:End Sub

 Sub sw53_Hit:Playsoundatball "sensor":Controller.Switch(53) = 1:End Sub
 Sub sw53_UnHit:Controller.Switch(53) = 0:End Sub

 ' Other switches

 Sub sw25_Hit:Controller.Switch(25) = 1:BallInShooterLane = 1:PlaySoundatball "sensor":End Sub
 Sub sw25_Unhit:Controller.Switch(25) = 0:BallInShooterLane = 0:End Sub
 Sub sw35_Hit
 Controller.Switch(35) = 1
 PlaySoundat "gateback_low",sw35
 'If ActiveBall.VelY <= -28 Then
 'PlaySoundat "subway2",sw35
 'End If
 xtype = 1
 ytype = 0
 End Sub


 Sub sw35_Unhit:Controller.Switch(35) = 0:End Sub
 Sub sw36_Hit:Controller.Switch(36) = 1:PlaySoundat "gateback_low",sw36:End Sub
 Sub sw36_Unhit:Controller.Switch(36) = 0:End Sub
 Sub sw75_Hit
 Controller.Switch(75) = 1
 PlaySoundat "gateback_low",sw75
 'If ActiveBall.VelY <= -28 Then
 'PlaySoundat "subway2",sw75
 'End If
 xtype = 1
 ytype = 0
 End Sub
 Sub sw75_Unhit:Controller.Switch(75) = 0:End Sub
 Sub sw76_Hit:Controller.Switch(76) = 1:PlaySoundat "gateback_low",sw76:End Sub
 Sub sw76_Unhit:Controller.Switch(76) = 0:End Sub


 'Targets

 Sub sw28_Hit:vpmTimer.PulseSw 28:PlaySoundat SoundFX("target",DOFTargets),sw28:End Sub
 Sub sw31_Hit:vpmTimer.PulseSw 31:PlaySoundat SoundFX("target",DOFTargets),sw31:End Sub
 Sub sw32_Hit:vpmTimer.PulseSw 32:PlaySoundat SoundFX("target",DOFTargets),sw32:End Sub
 Sub sw34_Hit:vpmTimer.PulseSw 34:PlaySoundat SoundFX("target",DOFTargets),sw34:End Sub
 Sub sw46_Hit:vpmTimer.PulseSw 46:PlaySoundat SoundFX("target",DOFTargets),sw46:End Sub
 Sub sw47_Hit:vpmTimer.PulseSw 47:PlaySoundat SoundFX("target",DOFTargets),sw47:End Sub
 Sub sw48_Hit:vpmTimer.PulseSw 48:PlaySoundat SoundFX("target",DOFTargets),sw48:End Sub
 Sub sw55_Hit:vpmTimer.PulseSw 55:PlaySoundat SoundFX("target",DOFTargets),sw55:End Sub
 Sub sw56_Hit:vpmTimer.PulseSw 56:PlaySoundat SoundFX("target",DOFTargets),Primitive29:End Sub
 Sub sw57_Hit:vpmTimer.PulseSw 57:PlaySoundat SoundFX("target",DOFTargets),Primitive28:End Sub
 Sub sw58_Hit:vpmTimer.PulseSw 58:PlaySoundat SoundFX("target",DOFTargets),Primitive28:End Sub

 '******
 ' vuks
 '******

 Dim popperTBall, popperTZpos
 Dim popperBBall, popperBZpos

 Sub sw61_Hit:PlaySoundat "kicker_enter",sw61:sw61a.Enabled=1:bsTBP.AddBall Me:End Sub

 Sub sw62_Hit:PlaySoundatvol "warehousehit",sw62, 0.2:sw62a.Enabled=1:bsBBP.AddBall Me:End Sub

 Sub SolTopPopper(Enabled)
     If Enabled Then
         If bsTBP.Balls Then
             sw61.Enabled = 0
             Set popperTBall = sw61a.Createball
             popperTBall.Z = 0
             popperTZpos = 0
             PlaySoundat SoundFX("popper",DOFContactors),sw61a
             sw61a.TimerInterval = 2
             sw61a.TimerEnabled = 1
         End If
     End If
 End Sub

 Sub sw61a_Timer
     popperTBall.Z = popperTZpos
     popperTZpos = popperTZpos + 10
     If popperTZpos> 170 Then
         sw61a.TimerEnabled = 0
         sw61a.DestroyBall
         bsTBP.ExitSol_On
     End If
 End Sub

 Sub SolBottomPopper(Enabled)
     If Enabled Then
         If bsBBP.Balls Then
             sw62.Enabled = 0
             Set popperBBall = sw62a.Createball
             popperBBall.Z = 0
             popperBZpos = 0
             PlaySoundat SoundFX("popper",DOFContactors),sw62a
             sw62a.TimerInterval = 2
             sw62a.TimerEnabled = 1
         End If
     End If
 End Sub

 Sub sw62a_Timer
     popperBBall.Z = popperBZpos
     popperBZpos = popperBZpos + 10
     If popperBZpos> 110 Then
         sw62a.TimerEnabled = 0
         sw62a.DestroyBall
         xtype = 0
         ytype = 1
         bsBBP.ExitSol_On
         If Sloppy_Turbo_VUK = 1 Then
         sw62.Enabled = 0 '' TEST
         else
         sw62.Enabled = 1
         End If
     End If
 End Sub

Dim tballcount
tballcount = 0

 Sub sw62b_Hit()
     sw62b.Destroyball
     '''sw62.Enabled = 1
     tballcount = tballcount + 1
     StopSound "rail"
     PlaySoundat "fx_metalclank",Turbo_Bottom
     vpmTimer.AddTimer 250, "AddBallToTurbo"
     If tballcount = 1 Then
	 ball1.visible = 1
     TREF.image = "disc2db_1"
	 End If
	 If tballcount = 2 Then
	 ball2.visible = 1
     TREF.image = "disc2db_2"
	 End If
	 If tballcount = 3 Then
	 ball3.visible = 1
     TREF.image = "disc2db_3"
	 End If
	 If tballcount = 4 Then
	 ball4.visible = 1
     TREF.image = "disc2db_4"
	 End If
 End Sub

 'Saucers

 Sub sw64_Hit:PlaySoundat "kicker_enter",sw64:bsUE.AddBall Me:End Sub
 Sub sw65_Hit:PlaySoundatvol "headquarterhit", sw65, 0.1:bsLE.AddBall Me:End Sub
 Sub Drain1_Hit:Playsoundatvol "drain",drain1,0.5:bsTrough.AddBall Me:End Sub
 'Sub Drain2_Hit:Playsoundatvol "drain",drain2,0.5:bsTrough.AddBall Me:End Sub
 'Sub Drain3_Hit:Playsoundatvol "drain",drain3,0.5:bsTrough.AddBall Me:End Sub
 'Sub Drain4_Hit:Playsoundatvol "drain",drain4,0.5:bsTrough.AddBall Me:End Sub
 'Sub Drain5_Hit:Playsoundatvol "drain",drain5,0.5:bsTrough.AddBall Me:End Sub

 ' Bumpers
 Sub Bumper1_Hit
 vpmTimer.PulseSw 72
 PlaySoundatvol SoundFX("bumper1",DOFContactors),bumper1, 1.0
 End Sub

 Sub Bumper2_Hit
 vpmTimer.PulseSw 73
 PlaySoundatvol SoundFX("bumper2",DOFContactors),bumper1, 1.0
 End Sub

 Sub Bumper3_Hit
 vpmTimer.PulseSw 74
 PlaySoundatvol SoundFX("bumper3",DOFContactors),bumper1, 1.0
 End Sub

 '************************************************************************
'					Slingshots Animation
'************************************************************************

Dim LStep, RStep

Sub LeftSlingShot_Slingshot
    PlaySoundat SoundFX("SlingshotLeft_X",DOFContactors),sling1
    vpmTimer.PulseSw 26
	LSling.Visible = 0
	LSling1.Visible = 1
	sling1.TransZ = -20
	LStep = 0
	LeftSlingShot.TimerEnabled = 1
    LeftSlingShot.TimerInterval = 10
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling1.TransZ = -10
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling1.TransZ = 0:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub


Sub RightSlingShot_Slingshot
	PlaySoundat SoundFX("SlingshotRight_X",DOFContactors),sling2
	vpmTimer.PulseSw 27
	RSling.Visible = 0
	RSling1.Visible = 1
	sling2.TransZ = -20
	RStep = 0
	RightSlingShot.TimerEnabled = 1
    RightSlingShot.TimerInterval = 10
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling2.TransZ = -10
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling2.TransZ = 0:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

 Sub sw54_Slingshot()
 vpmTimer.PulseSw 54
 playsoundat SoundFX("SlingShot",DOFContactors),BR3
 End Sub

 '***********
 ' Update GI
 '***********

 Set GiCallback2 = GetRef("UpdateGI")


 '***********
' Update GI
'***********
Dim xx
Dim gistep
gistep = 1/8

Sub UpdateGI(no, step)

If step < 5 Then
F27_LowBoy.State = 1
F28_LowBoy.State = 1
else
F27_LowBoy.State = 0
F28_LowBoy.State = 0
End If

If step => 0 Then
For each xx in GITL:xx.state = 1:next
For each xx in GITR:xx.state = 1:next
For each xx in GIB:xx.state = 1:next
End If

For each xx in GIF:xx.opacity = step * 30:next

If step = 0 then
Table1.ColorGradeImage = "LUT0"
else
if step = 1 Then
Table1.ColorGradeImage = "LUT0_1"
else
if step = 2 Then
Table1.ColorGradeImage = "LUT0_2"
else
if step = 3 Then
Table1.ColorGradeImage = "LUT0_3"
else
if step = 4 Then
Table1.ColorGradeImage = "LUT0_4"
else
if step = 5 Then
Table1.ColorGradeImage = "LUT0_5"
else
if step = 6 Then
Table1.ColorGradeImage = "LUT0_6"
else
if step = 7 Then
Table1.ColorGradeImage = "LUT0_7"
else
if step = 8 Then
Table1.ColorGradeImage = "LUT0_8"
End If
End If
End If
End If
End If
End If
End If
End If
End If

if step = 8 Then
	DOF 302, DOFOn
Else
	DOF 302, DOFOff
end if

    Select Case no

        Case 0 'top left

For each xx in GITL:xx.IntensityScale = gistep * step:next
ref5.opacity = step * 2
ref7.opacity = step * 2

If Dozer_Cab = 1 Then
If step = 1 then Controller.B2SSetData 201,0:Controller.B2SSetData 202,0:Controller.B2SSetData 203,0:Controller.B2SSetData 204,0:Controller.B2SSetData 205,0:Controller.B2SSetData 206,0
If step = 2 then Controller.B2SSetData 201,1:Controller.B2SSetData 202,0:Controller.B2SSetData 203,0:Controller.B2SSetData 204,0:Controller.B2SSetData 205,0:Controller.B2SSetData 206,0
If step = 3 then Controller.B2SSetData 202,1:Controller.B2SSetData 201,0:Controller.B2SSetData 203,0:Controller.B2SSetData 204,0:Controller.B2SSetData 205,0:Controller.B2SSetData 206,0
If step = 4 then Controller.B2SSetData 203,1:Controller.B2SSetData 202,0:Controller.B2SSetData 201,0:Controller.B2SSetData 204,0:Controller.B2SSetData 205,0:Controller.B2SSetData 206,0
If step = 5 then Controller.B2SSetData 204,1:Controller.B2SSetData 202,0:Controller.B2SSetData 203,0:Controller.B2SSetData 201,0:Controller.B2SSetData 205,0:Controller.B2SSetData 206,0
If step = 6 then Controller.B2SSetData 205,1:Controller.B2SSetData 202,0:Controller.B2SSetData 203,0:Controller.B2SSetData 204,0:Controller.B2SSetData 201,0:Controller.B2SSetData 206,0
If step = 8 then Controller.B2SSetData 206,1:Controller.B2SSetData 202,0:Controller.B2SSetData 203,0:Controller.B2SSetData 204,0:Controller.B2SSetData 205,0:Controller.B2SSetData 201,0
End If
        Case 1 'top right

For each xx in GITR:xx.IntensityScale = gistep * step:next
ref6.opacity = step * 2

If Dozer_Cab = 1 Then
If step = 1 then Controller.B2SSetData 101,0:Controller.B2SSetData 102,0:Controller.B2SSetData 103,0:Controller.B2SSetData 104,0:Controller.B2SSetData 105,0:Controller.B2SSetData 106,0
If step = 2 then Controller.B2SSetData 101,1:Controller.B2SSetData 102,0:Controller.B2SSetData 103,0:Controller.B2SSetData 104,0:Controller.B2SSetData 105,0:Controller.B2SSetData 106,0
If step = 3 then Controller.B2SSetData 102,1:Controller.B2SSetData 101,0:Controller.B2SSetData 103,0:Controller.B2SSetData 104,0:Controller.B2SSetData 105,0:Controller.B2SSetData 106,0
If step = 4 then Controller.B2SSetData 103,1:Controller.B2SSetData 102,0:Controller.B2SSetData 101,0:Controller.B2SSetData 104,0:Controller.B2SSetData 105,0:Controller.B2SSetData 106,0
If step = 5 then Controller.B2SSetData 104,1:Controller.B2SSetData 102,0:Controller.B2SSetData 103,0:Controller.B2SSetData 101,0:Controller.B2SSetData 105,0:Controller.B2SSetData 106,0
If step = 6 then Controller.B2SSetData 105,1:Controller.B2SSetData 102,0:Controller.B2SSetData 103,0:Controller.B2SSetData 104,0:Controller.B2SSetData 101,0:Controller.B2SSetData 106,0
If step = 8 then Controller.B2SSetData 106,1:Controller.B2SSetData 102,0:Controller.B2SSetData 103,0:Controller.B2SSetData 104,0:Controller.B2SSetData 105,0:Controller.B2SSetData 101,0
End If
        Case 2 'bottom

For each xx in GIB:xx.IntensityScale = gistep * step:next
ref1.opacity = step * 2
ref2.opacity = step * 2
ref3.opacity = step * 2
ref4.opacity = step * 2

If Dozer_Cab = 1 Then
If step = 1 then Controller.B2SSetData 111,0:Controller.B2SSetData 112,0:Controller.B2SSetData 113,0:Controller.B2SSetData 114,0:Controller.B2SSetData 115,0:Controller.B2SSetData 116,0
If step = 2 then Controller.B2SSetData 111,1:Controller.B2SSetData 112,0:Controller.B2SSetData 113,0:Controller.B2SSetData 114,0:Controller.B2SSetData 115,0:Controller.B2SSetData 116,0
If step = 3 then Controller.B2SSetData 112,1:Controller.B2SSetData 111,0:Controller.B2SSetData 113,0:Controller.B2SSetData 114,0:Controller.B2SSetData 115,0:Controller.B2SSetData 116,0
If step = 4 then Controller.B2SSetData 113,1:Controller.B2SSetData 112,0:Controller.B2SSetData 111,0:Controller.B2SSetData 114,0:Controller.B2SSetData 115,0:Controller.B2SSetData 116,0
If step = 5 then Controller.B2SSetData 114,1:Controller.B2SSetData 112,0:Controller.B2SSetData 113,0:Controller.B2SSetData 111,0:Controller.B2SSetData 115,0:Controller.B2SSetData 116,0
If step = 6 then Controller.B2SSetData 115,1:Controller.B2SSetData 112,0:Controller.B2SSetData 113,0:Controller.B2SSetData 114,0:Controller.B2SSetData 111,0:Controller.B2SSetData 116,0
If step = 8 then Controller.B2SSetData 116,1:Controller.B2SSetData 112,0:Controller.B2SSetData 113,0:Controller.B2SSetData 114,0:Controller.B2SSetData 115,0:Controller.B2SSetData 111,0
End If

   End Select
End Sub


'**********
 ' TopCar
 '**********

 Sub SolTopCar(Enabled)
     If Enabled Then
         PlaySoundat SoundFX("motor",DOFGear),Indy_Top
         Indy_Upper.Enabled = 1
         Indy_up_Flash.Enabled = 1
     Else
         Indy_Upper.Enabled = 0
         Indy_up_Flash.Enabled = 0
		 XBallShadow1.visible = 0
		 IU1.State = 0
		 IU2.State = 0
		 IU3.State = 0
		 IU4.State = 0
		 IU5.State = 0
		 IUFLASH1.visible = 0
         IUFLASH2.visible = 0
		 Car_Shadow.visible = 1
         F17s4.visible = 0
     End If
 End Sub

Sub Indy_Upper_Timer()
indy_top.objrotz = indy_top.objrotz + 1
indy_shaft.objrotz = indy_shaft.objrotz + 1
topcar_base.objrotz = topcar_base.objrotz + 1
topcar_base1.objrotz = topcar_base1.objrotz + 1
If Rotating_Track = 1 Then
track.objrotz = track.objrotz -1
End If
car_shadow.objrotz = car_shadow.objrotz + 1
XBallShadow1.objrotz = XBallShadow1.objrotz + 1
End Sub

Dim iupflash

iupflash = 1

Sub Indy_Up_Flash_Timer
Select Case iupflash
Case 1:
XBallShadow1.visible = 0
IU1.State = 0
IU2.State = 0
IU3.State = 0
IU4.State = 0
IU5.State = 0
IUFLASH1.visible = 0
IUFLASH2.visible = 0
Car_Shadow.visible = 0
F17s4.visible = 0
iupflash = 2
Case 2
XBallShadow1.visible = 1
IU1.State = 1
IU2.State = 1
IU3.State = 1
IU4.State = 1
IU5.State = 1
IUFLASH1.visible = 1
IUFLASH2.visible = 1
F17s4.visible = 1
Car_Shadow.visible = 1
iupflash = 1
End Select
End Sub



 '*********************************************************************************************
 ' Turbo Handler Code by Dorsola - Code Additions for Primitive Objects Turbo (Dozer).
 ' Theory of operation:
 ' - TurboPopper    shoots a ball into the turbo mechanism
 ' - TurboIndex     gets triggered every quarter turn of the turbo mech
 ' - TurboBallSense detects the quadrant slats every quarter turn and a locked ball
 '
 ' * If storing balls, turbo turns at low speed and and keeps balls in turbo
 ' * If releasing balls, turbo spins up until centrIfugal force ejects each ball onto the ramp.
 '
 ' Known Issues:
 ' * Turbo lock behavior is not always consistent, but is believed to be accurately simulated

 '**********************************************************************************************

 Dim BallInTurboIndex(4), TurboSpeed, TurboMotorState, TurboRotatePosition, TurboCounter, TurboTargetCounter

 Sub CloseBallSense:Controller.Switch(63) = 1:End Sub

 Sub OpenBallSense:Controller.Switch(63) = 0:End Sub

Sub CloseTurboIndex:Controller.Switch(66) = 1:End Sub
 Sub OpenTurboIndex:Controller.Switch(66) = 0:End Sub

 Sub ShowTurboBalls
     ShowTurboPos TurboPos1A, BallInTurboIndex(1) And TurboRotatePosition = 0
     ShowTurboPos TurboPos1B, BallInTurboIndex(1) And TurboRotatePosition = 1
     ShowTurboPos TurboPos1C, BallInTurboIndex(1) And TurboRotatePosition = 2
     ShowTurboPos TurboPos1D, BallInTurboIndex(1) And TurboRotatePosition = 3
     ShowTurboPos TurboPos2A, BallInTurboIndex(2) And TurboRotatePosition = 0
     ShowTurboPos TurboPos2B, BallInTurboIndex(2) And TurboRotatePosition = 1
     ShowTurboPos TurboPos2C, BallInTurboIndex(2) And TurboRotatePosition = 2
     ShowTurboPos TurboPos2D, BallInTurboIndex(2) And TurboRotatePosition = 3
     ShowTurboPos TurboPos3A, BallInTurboIndex(3) And TurboRotatePosition = 0
     ShowTurboPos TurboPos3B, BallInTurboIndex(3) And TurboRotatePosition = 1
     ShowTurboPos TurboPos3C, BallInTurboIndex(3) And TurboRotatePosition = 2
     ShowTurboPos TurboPos3D, BallInTurboIndex(3) And TurboRotatePosition = 3
     ShowTurboPos TurboPos4A, BallInTurboIndex(4) And TurboRotatePosition = 0
     ShowTurboPos TurboPos4B, BallInTurboIndex(4) And TurboRotatePosition = 1
     ShowTurboPos TurboPos4C, BallInTurboIndex(4) And TurboRotatePosition = 2
     ShowTurboPos TurboPos4D, BallInTurboIndex(4) And TurboRotatePosition = 3
 End Sub

 Sub ShowTurboWalls
     Select Case TurboRotatePosition
         Case 0:TurboWallA.isdropped = 0:TurboWallB.isdropped = 1:TurboWallC.isdropped = 1:TurboWallD.isdropped = 1
         Case 1:TurboWallA.isdropped = 1:TurboWallB.isdropped = 0:TurboWallC.isdropped = 1:TurboWallD.isdropped = 1
         Case 2:TurboWallA.isdropped = 1:TurboWallB.isdropped = 1:TurboWallC.isdropped = 0:TurboWallD.isdropped = 1
         Case 3:TurboWallA.isdropped = 1:TurboWallB.isdropped = 1:TurboWallC.isdropped = 1:TurboWallD.isdropped = 0
     End Select
 End Sub

'Dozer Code Mod
 Sub ShowTurboPos(kicker, show)
     If show Then
         If kicker.Enabled = 0 Then
             'kicker.CreateBall
             kicker.Enabled = 1
         End If
     Else
         If kicker.Enabled Then
             'kicker.DestroyBall
             kicker.Enabled = 0
         End If
     End If
 End Sub
'Dozer Code Mod

 Sub Init_Turbo()
     Dim count
     for count = 1 to 4:BallInTurboIndex(count) = 0:next
     TurboSpeed = 0
     TurboCounter = 0
     TurboRotatePosition = 0
     TurboMotorState = Controller.GetMech(0)
     OpenTurboIndex
     OpenBallSense
     ShowTurboBalls
     ShowTurboWalls
 End Sub

 Sub AdvanceTurboBalls()
     Dim overflow
     overflow = BallInTurboIndex(4)
     Dim count
     for count = 3 to 1 step - 1
         BallInTurboIndex(count + 1) = BallInTurboIndex(count)
     next

     BallInTurboIndex(1) = overflow
 End Sub

 Sub RotateTurbo()
     TurboRotatePosition = (TurboRotatePosition + 1) mod 4
     Select Case TurboRotatePosition
         Case 0
             AdvanceTurboBalls
             OpenTurboIndex
             If BallInTurboIndex(2) Then
                 CloseBallSense
             Else
                 OpenBallSense
             End If

             If TurboSpeed> 0 and TurboMotorState = 0 Then TurboSpeed = 0
         Case 1
             CloseTurboIndex
             CloseBallSense
         Case 2
             CloseTurboIndex
             CloseBallSense
             If TurboSpeed = 2 And BallInTurboIndex(2) Then
				TurboPopperEject.Enabled = 1
                 TurboPopperEject.CreateBall
                 TurboPopperEject.Kick 80, 26
                 tballcount = tballcount - 1
                 PlaySoundat "metalhit",TurboPopperEject
                 ball1.visible = 0
                 ball2.visible = 0
                 ball3.visible = 0
                 ball4.visible = 0
                 TREF.image = "disc2db"
                 BallInTurboIndex(2) = 0
             End If
         Case 3
             CloseTurboIndex
             CloseBallSense
     End Select

     ShowTurboBalls
     ShowTurboWalls
 End Sub

 Sub AddBallToTurbo(no)
     Dim index
     index = 1
     If BallInTurboIndex(1) Then
         While BallInTurboIndex(index) = 1 and index <= 4
             index = index + 1
         WEnd
     End If

     If index> 4 Then
         sw62b.CreateBall
         sw62b_Hit
     Else
         BallInTurboIndex(index) = 1
         If index = 2 Then CloseBallSense
     End If

     ShowTurboBalls
     ShowTurboWalls
 End Sub

 Sub TurboMasterTimer_Timer()
     Dim Temp
     Temp = Controller.GetMech(0)
     If Temp <> TurboMotorState Then
         TurboMotorState = Temp
         If Temp> 0 Then TurboSpeed = Temp
         'Dozer Code
         If Temp = 1 Then
         Turbo_Slow.enabled = 1
         PlaySoundatvol SoundFX("TurboMotor_Low",DOFGear),Turbo_Bottom, 0.05
         'PlaySound "TurboMotor_Low", 1, 0.2, AudioPan(Turbo_Bottom), 0,0,0, 1, AudioFade(Turbo_Bottom)
		 DOF 303, DOFOn
         Turbo_Fast.enabled = 0:StopSound "TurboMotor":DOF 301, DOFOff
         Turbo_Flash_fast.enabled = 0
         Turbo_Flash_slow.enabled = 1
         else
         If Temp = 2 Then
         Turbo_Slow.enabled = 0:StopSound "TurboMotor_Low":DOF 303, DOFOff
         Turbo_Fast.enabled = 1
         PlaySoundatvol SoundFX("TurboMotor",DOFGear),Turbo_Bottom, 0.05
         DOF 301, DOFOn
         Turbo_Flash_fast.enabled = 1
         Turbo_Flash_slow.enabled = 0
         else
         Turbo_Fast.enabled = 0
         Turbo_Slow.enabled = 0
         trepos=1:Turbo_Repos.enabled = 1
         End If
         End If
         'Dozer Code
      End If

     Select Case TurboSpeed
         Case 0
             TurboCounter = 0
         Case 1
             TurboCounter = TurboCounter + 1
             If TurboCounter> 2 Then
                 RotateTurbo
                 TurboCounter = 0
             End If
         Case 2
             RotateTurbo
             TurboCounter = 0
     End Select
 End Sub


' Code to handle Primitive Turbo and Ball Handling (Dozer).

Dim tfpos,trepos
tfpos = 60

Sub Turbo_Fast_timer()
If tfpos = 360 Then
tfpos = 0
End If
impellor.objrotz = impellor.objrotz + 1
tref.rotz = tref.rotz - 1
modref.rotz = modref.rotz - 1
turbo_bottom.objrotz = turbo_bottom.objrotz + 1
ball1.objrotz = ball1.objrotz + 1
ball2.objrotz = ball2.objrotz + 1
ball3.objrotz = ball3.objrotz + 1
ball4.objrotz = ball4.objrotz + 1
tfpos = tfpos + 1
end Sub

Dim tflashf
tflashf = 1

Sub Turbo_Flash_Fast_Timer
Select Case tflashf
Case 1:
Fturbo.visible = 0:Turbo_Side_Flash2.visible = 0:tflashf = 2
Case 2
Fturbo.visible = 1:Turbo_Side_Flash2.visible = 1:tflashf = 1
End Select
End Sub

Dim tflashs
tflashs = 1
Sub Turbo_Flash_Slow_Timer
Select Case tflashs
Case 1:
Fturbo.visible = 0:Turbo_Side_Flash2.visible = 0:tflashs = 2
Case 2
Fturbo.visible = 1:Turbo_Side_Flash2.visible = 1:tflashs = 1
End Select
End Sub

Sub Turbo_Slow_timer()
If tfpos = 360 Then
tfpos = 0
End If
impellor.objrotz = impellor.objrotz + 1
tref.rotz = tref.rotz - 1
modref.rotz = modref.rotz - 1
turbo_bottom.objrotz = turbo_bottom.objrotz + 1
ball1.objrotz = ball1.objrotz + 1
ball2.objrotz = ball2.objrotz + 1
ball3.objrotz = ball3.objrotz + 1
ball4.objrotz = ball4.objrotz + 1
tfpos = tfpos + 1
end Sub

Sub Turbo_Repos_Timer()
Select Case trepos
Case 1:
If NOT tfpos = 60 Then
trepos = 2
End If
Case 2:
If tfpos => 60 Then
impellor.objrotz = 60
tfpos = 60
Fturbo.visible = 1
turbo_flash_slow.enabled = 0
turbo_flash_fast.enabled = 0
StopSound "TurboMotor_Low":DOF 303, DOFOff
StopSound "TurboMotor":DOF 301, DOFOff
vpmTimer.AddTimer 200, "RestopSound"
me.enabled = 0
End If
If tfpos = 360 Then
tfpos = 0
End If
impellor.objrotz = impellor.objrotz + 1
tref.rotz = tref.rotz - 1
modref.rotz = modref.rotz - 1
turbo_bottom.objrotz = turbo_bottom.objrotz + 1
ball1.objrotz = ball1.objrotz + 1
ball2.objrotz = ball2.objrotz + 1
ball3.objrotz = ball3.objrotz + 1
ball4.objrotz = ball4.objrotz + 1
tfpos = tfpos + 1
End Select
End Sub

Sub RestopSound(no)
StopSound "TurboMotor_Low":DOF 303, DOFOff
StopSound "TurboMotor":DOF 301, DOFOff
End Sub

'Redundant but kept until release.
Sub turbo_watch_timer()
If tballcount = 1 Then
ball1.visible = 1
End If
If tballcount = 2 Then
ball2.visible = 1
End If
If tballcount = 3 Then
ball3.visible = 1
End If
If tballcount = 4 Then
ball4.visible = 1
End If
End Sub


'Insert Lights

Set Lights(11)=L11
Set Lights(12)=L12
Set Lights(13)=L13
Set Lights(15)=L15
Set Lights(18)=L18
'Set Lights(14)=L14
'Set Lights(16)=L16
'Set Lights(17)=L17
Lights(14)=Array(L14,L14a)
Lights(16)=Array(L16,L16a)
Lights(17)=Array(L17,L17a)
'Set Lights(21)=L21
'Set Lights(22)=L22
'Set Lights(23)=L23
Lights(21)=Array(L21,L21a)
Lights(22)=Array(L22,L22a)
Lights(23)=Array(L23,L23a)
Set Lights(24)=L24
Set Lights(25)=L25
Set Lights(26)=L26
Set Lights(27)=L27
Set Lights(28)=L28
Set Lights(31)=L31
Set Lights(32)=L32
Set Lights(33)=L33
Set Lights(34)=L34
Set Lights(35)=L35
Set Lights(36)=L36
Set Lights(37)=L37
Set Lights(38)=L38
Set Lights(41)=L41
Set Lights(42)=L42
Set Lights(43)=L43
Set Lights(44)=L44
Set Lights(45)=L45
Set Lights(46)=L46
Set Lights(47)=L47
Set Lights(48)=L48
Set Lights(57)=L57
Set Lights(58)=L58
Set Lights(61)=L61
Set Lights(62)=L62
Set Lights(63)=L63
Set Lights(51)=L51
Set Lights(52)=L52
Set Lights(53)=L53
Set Lights(54)=L54
Set Lights(55)=L55
Set Lights(56)=L56
Set Lights(64)=L64
Set Lights(65)=L65
Set Lights(66)=L66
Set Lights(75)=L75x
Set Lights(76)=L76x
Set Lights(77)=L77x
Set Lights(78)=L78x
Lights(71)=Array(L71x,L71)
Lights(72)=Array(L72x,L72)
Lights(73)=Array(L73x,L73)
Lights(74)=Array(L74x,L74)
Lights(75)=Array(L75x,L75x1)
Lights(76)=Array(L76x,L76x1)
Lights(77)=Array(L77x,L77x1)
Lights(78)=Array(L78x,L78x1)
Lights(81)=Array(L81x,L81x1)
Lights(82)=Array(L82x,L82x1)
Lights(83)=Array(L83x,L83x1)
Lights(84)=Array(L84x,L84x1)


 '----------------------------
 ' Misc Hitty Stuff
 '----------------------------

 Sub TurboErrorCatcher_Hit()
     Me.DestroyBall
	TurboPopperEject.Enabled = 1
     TurboPopperEject.CreateBall
     TurboPopperEject.Kick 80, 26
     tballcount = tballcount - 1
 End Sub

Sub TurboPopperEject_Unhit() : TurboPopperEject.Enabled = 0 : End Sub
Sub sw61a_UnHit() : PlaySoundat "balldrop",sw61a : xtype = 1:ytype = 0 : sw61a.Enabled=0 : End Sub
Sub sw62a_UnHit() : sw62a.Enabled=0 : End Sub
Sub sw62b_UnHit() : End Sub

Sub TOUT_Hit()
'PlaySoundat "rail",tout
ytype = 1
xtype = 0
End Sub

Sub Feed_Hit()
ytype = 1
End Sub

Sub RHD_Hit()
'PlaySound "balldrop",0,1,1.0,0
PlaySoundatvol "balldrop",RHD2, 0.5
'StopSound "subway2"
End Sub

Sub RHD2_Hit()
'StopSound "rail"
PlaySoundatvol "balldrop",RHD2, 0.5
'PlaySound "balldrop",0,1,1.0,0
End Sub

Sub LHD_Hit()
PlaySoundatvol "balldrop",sw16, 0.5
'StopSound "subway2"
End Sub

Sub MetalStop_Hit()
PlaySoundat "fx_metalclank",Primitive61
'StopSound "rail_low_slower"
End Sub

Sub Wall720_Hit()
PlaySoundat "metalhit",sw65
End Sub

Sub Wall14_Hit()
PlaySoundat "metalhit",whramp1
End Sub

Sub Wall205_Hit()
PlaySoundat "metalhit",sw47
End Sub

Sub Wall27_Hit()
PlaySoundat "metalhit",disc2
End Sub

Sub DiverterHold_Hit()
PlaySoundat "plastichit",Diverterhold
End Sub

Sub gate2_hit()
PlaySoundat "gateback_low",gate2
End Sub

Sub gate3_hit()
PlaySoundat "gateback_low",gate3
End Sub

Sub Tback_Hit()
If Sloppy_Turbo_VUK = 1 Then
PlaySoundat "metalhit",whramp1
mSpinMagnet.MagnetOn = True
tppos1 = 1:tpopenable.enabled = 1
End If
End Sub

Dim tppos1

Sub tpopenable_timer()
Select Case tppos1
Case 1:
mSpinMagnet.MagnetOn = False
sw62.enabled = 1
tppos1 = 2
Case 2:
me.enabled = 0
End Select
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

Sub PlaySoundAtVol(soundname, tableobj, vol)
    PlaySound soundname, 1, (vol), AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

Sub PlaySoundAtBall(soundname)
    PlaySoundAt soundname, ActiveBall
End Sub

'*****************************************
'      JP's VP10 Rolling Sounds
'*****************************************
Dim xtype, ytype

Const tnob = 12 ' total number of balls
ReDim rolling(tnob)
Redim xrolling(tnob)
Redim yrolling(tnob)
InitRolling
InitRolling2
InitRolling3

Sub InitRolling
    Dim i
    For i = 0 to tnob
        rolling(i) = False
    Next
End Sub

Sub InitRolling2
    Dim i
    For i = 0 to tnob
        xrolling(i) = False
    Next
End Sub

Sub InitRolling3
    Dim i
    For i = 0 to tnob
        yrolling(i) = False
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
            PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b)), AudioPan(BOT(b)), 0, Pitch(BOT(b)), 1, 0, AudioFade(BOT(b))
        Else
            If rolling(b) = True Then
                StopSound("fx_ballrolling" & b)
                rolling(b) = False
            End If
        End If

            If BallVel(BOT(b) ) > 1 AND BOT(b).z > 30 AND xtype = 1 Then
            xrolling(b) = True
            PlaySound("subway2"), -1, Vol(BOT(b)), AudioPan(BOT(b)), 0, Pitch(BOT(b)), 1, 0, AudioFade(BOT(b))
        Else
            If xrolling(b) = True Then
                StopSound("subway2")
                xrolling(b) = False
            End If
        End If

            If BallVel(BOT(b) ) > 1 AND BOT(b).z > 30 AND ytype = 1 Then
            yrolling(b) = True
            PlaySound("rail"), -1, Vol(BOT(b)), AudioPan(BOT(b)), 0, Pitch(BOT(b)), 1, 0, AudioFade(BOT(b))
        Else
            If yrolling(b) = True Then
                StopSound("rail")
                yrolling(b) = False
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
'      Other Table Sounds
'*****************************************

Sub aRubbers_Hit(idx)
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

Sub RightFlipper2_Collide(parm)
 	RandomSoundFlipper()
End Sub

Sub RandomSoundFlipper()
	Select Case Int(Rnd*3)+1
		Case 1 : PlaySound "flip_hit_1", 0, Vol(ActiveBall)*3, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 2 : PlaySound "flip_hit_2", 0, Vol(ActiveBall)*3, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 3 : PlaySound "flip_hit_3", 0, Vol(ActiveBall)*3, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End Select
End Sub

Sub Dampen(dt,df,r)						'dt is threshold speed, df is dampen factor 0 to 1 (higher more dampening), r is randomness
	Dim dfRandomness
	r=cint(r)
	dfRandomness=INT(RND*(2*r+1))
	df=df+(r-dfRandomness)*.01
	If ABS(activeball.velx) > dt Then activeball.velx=activeball.velx*(1-df*(ABS(activeball.velx)/100))
	If ABS(activeball.vely) > dt Then activeball.vely=activeball.vely*(1-df*(ABS(activeball.vely)/100))
End Sub

'*****************************************
'	Ball Shadow
'*****************************************

Dim BallShadow
BallShadow = Array (BallShadow1, BallShadow2, BallShadow3, BallShadow4, BallShadow5)


Sub BallShadowUpdate()
    Dim BOT, b
    BOT = GetBalls

	' render the shadow for each ball
    For b = 0 to UBound(BOT)
		If BOT(b).X < Table1.Width/2 Then
			BallShadow(b).X = ((BOT(b).X) - (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/7)) + 10
		Else
			BallShadow(b).X = ((BOT(b).X) + (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/7)) - 10
		End If

			BallShadow(b).Y = BOT(b).Y + 10
			BallShadow(b).Z = 1
		If BOT(b).Z > 20 Then
			BallShadow(b).visible = 1
		Else
			BallShadow(b).visible = 0
		End If
	Next
End Sub

Sub Table1_Exit
Controller.Stop
End Sub

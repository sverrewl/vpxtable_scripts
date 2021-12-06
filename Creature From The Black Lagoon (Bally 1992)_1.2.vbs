
'************************************************************
'************************************************************
'
'  Creature from the Black Lagoon (Bally 1992) / IPD No. 588 / 4 Players
'
'   Credits:
'
'      VPX by       fuzzel, flupper1
'      Script by    flupper, rothbauerw
'      Primitives by  fuzzel, flupper1
'    Supported by     randr, clarkkent, hauntfreaks
'
'************************************************************
'************************************************************

Option Explicit
Randomize

'******************************************************
'             OPTIONS
'******************************************************

' Thalamus 2018-08-06
' Table has its own "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Changed UseSolenoids=1 to 2
' No special SSF tweaks yet.

Const EnableBallControl = true  'set to false to disable the C key from taking manual control
Const VolumeDial = 5        'Change bolume of hit events
Const RollingSoundFactor = 3    'set sound level factor here for Ball Rolling Sound, 1=default level

dim chaselightson, chaselightsintensity, chaselightsbloom, indirectlightsintensity, shadowopacity, bumperslit, fslutfading, dtlutfading

chaselightsintensity = 0.5    ' usable range is 0.5 - 1
chaselightsbloom = 1      ' 0 or 1 (0 for less bloom, 1 for more bloom)
indirectlightsintensity = 1   ' usable range is 0.5 - 1
shadowopacity = 90        ' usable range is 0 - 100 (90 default)
bumperslit = 0          ' 0 or 1 (0 bumpers without lights (original), 1 bumpers with lights
fslutfading = 1         ' 1 enables GI lut fading for full screen, 0 to disable
dtlutfading = 0           ' 1 enables GI lut fading for desk top, 0 to disable

'******************************************************
'           STANDARD DEFINITIONS
'******************************************************

Dim Ballsize,BallMass
Ballsize = 50
Ballmass = 1.7

Const UseSolenoids = 2
' Thal : Added because of useSolenoid=2
Const cSingleLFlip = 0
Const cSingleRFlip = 0
Const UseLamps = 0
Const UseSync = 0
Const UseGI = 1
Const UseVPMModSol = 1

'******************************************************
'           TABLE INIT
'******************************************************

if Version < 10400 then msgbox "This table requires Visual Pinball 10.4 beta or newer!" & vbnewline & "Your version: " & Version/1000

On Error Resume Next
  ExecuteGlobal GetTextFile("controller.vbs")
  If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

LoadVPM "01560000", "WPC.VBS", 3.50
Const cGameName="cftbl_l4"

dim bsSubway
Sub Table1_Init
  vpmInit Me
  With Controller
    .GameName = cGameName
    If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
    .SplashInfoLine = "Creature from the Black Lagoon (Bally 1992)"
    .HandleMechanics=0
    .HandleKeyboard=0
    .ShowDMDOnly=1
    .ShowFrame=0
    .ShowTitle=0
        '.hidden = not table1.showdt
         On Error Resume Next
         .Run GetPlayerHWnd
         If Err Then MsgBox Err.Description
         On Error Goto 0
     End With

  dim light
  If chaselightsbloom = 0 Then
    chaselightsintensity = chaselightsintensity * 8
    For Each Light in chaselights: Light.Falloff=20:Light.Falloffpower=8:Light.IntensityScale = chaselightsintensity : Light.FadeSpeedUp = Light.IntensityScale * Light.Intensity / 20 : Light.FadeSpeedDown = Light.FadeSpeedUp / 3: Next
  Else
    For Each Light in chaselights: Light.IntensityScale = chaselightsintensity : Light.FadeSpeedUp = Light.IntensityScale * Light.Intensity / 20 : Light.FadeSpeedDown = Light.FadeSpeedUp / 3: Next
  End If

  For Each Light in indirectlights: Light.IntensityScale = indirectlightsintensity : Light.FadeSpeedUp = Light.IntensityScale * Light.Intensity / 60 : Light.FadeSpeedDown = Light.FadeSpeedUp : Next

  Flasher21.opacity = shadowopacity

    '************  Main Timer init  ********************

  PinMAMETimer.Interval = PinMAMEInterval
  PinMAMETimer.Enabled = 1

    '************  Nudging   **************************

  vpmNudge.TiltSwitch = 14
  vpmNudge.Sensitivity = 5
  'Add Slings - WJR
  vpmNudge.TiltObj = Array(Bumper1, Bumper2, Bumper3, LeftSlingShot, RightSlingShot)

    '************  Trough **************************
  sw58.CreateSizedballWithMass Ballsize/2,Ballmass
  sw57.CreateSizedballWithMass Ballsize/2,Ballmass
  sw56.CreateSizedballWithMass Ballsize/2,Ballmass

  Controller.Switch(58) = 1
  Controller.Switch(57) = 1
  Controller.Switch(56) = 1

  Controller.Switch(15) = 1

end sub


'******************************************************
'             KEYS
'******************************************************

Sub Table1_KeyDown(ByVal keycode)
  If keycode = PlungerKey Then
    Plunger.PullBack
    PlaySound "plungerpull",0,1,AudioPan(Plunger),0.25,0,0,1,AudioFade(Plunger)
  End If

  If keycode = LeftTiltKey Then
    Nudge 90, 2
  End If

  If keycode = RightTiltKey Then
    Nudge 270, 2
  End If

  If keycode = CenterTiltKey Then
    Nudge 0, 2
  End If

  '************************   Start Ball Control 1/3
  if keycode = 46 then        ' C Key
    If contball = 1 Then
      contball = 0
    Else
      contball = 1
    End If
  End If
  if keycode = 48 then        'B Key
    If bcboost = 1 Then
      bcboost = bcboostmulti
    Else
      bcboost = 1
    End If
  End If
  if keycode = 203 then bcleft = 1    ' Left Arrow
  if keycode = 200 then bcup = 1      ' Up Arrow
  if keycode = 208 then bcdown = 1    ' Down Arrow
  if keycode = 205 then bcright = 1   ' Right Arrow
  '************************   End Ball Control 1/3

  If vpmKeyDown(keycode) Then Exit Sub
End Sub

Sub Table1_KeyUp(ByVal keycode)
  If keycode = PlungerKey Then
    Plunger.Fire
    PlaySound "plunger",0,1,AudioPan(Plunger),0.25,0,0,1,AudioFade(Plunger)
  End If

  '************************   Start Ball Control 2/3
  'if keycode = 203 then bcleft = 0   ' Left Arrow
  'if keycode = 200 then bcup = 0     ' Up Arrow
  'if keycode = 208 then bcdown = 0   ' Down Arrow
  'if keycode = 205 then bcright = 0    ' Right Arrow
  '************************   End Ball Control 2/3

  If vpmKeyUp(keycode) Then Exit Sub

End Sub

'******************************************************
'             SOLENOIDS
'******************************************************

SolCallback(1) =  "RightUpperKicker"
SolCallback(3) =  "RightLowerKicker"
SolCallback(4) =  "ReleaseBall"
SolCallback(12) =   "SolOuthole"


'********Solenoids********

'  SolCallBack(1) = "TopRightPopper"
'  SolCallBack(2) = "SetLamp 103,"
'  SolCallBack(3) = "bsSubway.SolOut"
'  SolCallBack(4) = "bsTrough.SolOut"
SolCallBack(7) = "vpmSolSound SoundFX(""Knocker"",DOFKnocker),"
'  SolCallBack(8) = "SetLamp 106,"
'  SolCallBack(9) = "BackFlash"
'  SolCallBack(10) = "SetLamp 100,"
'  SolCallBack(11) = "CreatureFlash"
'  SolCallBack(12) = "bsTrough.SolIn"
'  SolCallBack(16) = "SetLamp 109,"
'  SolCallBack(17) = "SetLamp 110,"
'  SolCallBack(18) = "SetLamp 101,"
'  SolCallBack(19) = "SetLamp 102,"
' ' SolCallBack(20) = "SeqGI1"
' ' SolCallBack(21) = "HologramPushMotor"
'  SolCallBack(22) = "SetLamp 111,"
'  SolCallBack(23) = "DropRUp"
' ' SolCallBack(24) = "SeqGI2"
'  SolCallBack(25) = "StartMov"
'  SolCallBack(26) = "DropRDown"
'  SolCallBack(27) = "CreatureMotorMirror"
'  SolCallBack(28) = "HoloLamp"
'
' '******Flipper Subs
'
SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

'does not work properly with rom !!! so I did not use this, but for reference

'dim val1, val2,
'
'SolCallBack(20) = "SeqGI1"
'SolCallBack(24) = "SeqGI2"
'
'Sub SeqGI1(Enabled)
'    If Enabled Then val1 = 1 Else val1 = 0 end if
' chaseval = val1 + val2
'End Sub
'
'Sub SeqGI2(Enabled)
'    If Enabled Then val2 = 2 Else val2 = 0 end if
' chaseval = val1 + val2
'End Sub

dim chaseval

Sub swirlramplight1_Timer()
  dim chase1, chase2, chase3, chase4
  chaseval = chaseval + 1
  If chaseval > 3 then chaseval = 0 end if
  If chaselightson Then
    Select Case chaseval
      Case 0 : chase1 = 0 : chase2 = 1 : chase3 = 1 : chase4 = 1
      Case 1 : chase1 = 1 : chase2 = 0 : chase3 = 1 : chase4 = 1
      Case 2 : chase1 = 1 : chase2 = 1 : chase3 = 0 : chase4 = 1
      Case 3 : chase1 = 1 : chase2 = 1 : chase3 = 1 : chase4 = 0
    End Select
  Else
    chase1 = 0 : chase2 = 0 : chase3 = 0 : chase4 = 0
  End if
  swirlramplight1.state = chase1 : swirlramplight5.state = chase1 : swirlramplight9.state = chase1 : swirlramplight13.state = chase1 : swirlramplight17.state = chase1 : swirlramplight21.state = chase1 : swirlramplight25.state = chase1 : swirlramplight29.state = chase1
  swirlramplight2.state = chase2 : swirlramplight6.state = chase2 : swirlramplight10.state = chase2 : swirlramplight14.state = chase2 : swirlramplight18.state = chase2 : swirlramplight22.state = chase2 : swirlramplight26.state = chase2 : swirlramplight30.state = chase2
  swirlramplight3.state = chase3 : swirlramplight7.state = chase3 : swirlramplight11.state = chase3 : swirlramplight15.state = chase3 : swirlramplight19.state = chase3 : swirlramplight23.state = chase3 : swirlramplight27.state = chase3 : swirlramplight31.state = chase3
  swirlramplight4.state = chase4 : swirlramplight8.state = chase4 : swirlramplight12.state = chase4 : swirlramplight16.state = chase4 : swirlramplight20.state = chase4 : swirlramplight24.state = chase4 : swirlramplight28.state = chase4 : swirlramplight32.state = chase4
  if chase2=0 or swirlramplight2.intensityscale < 0.01 Then
    wflash6.intensityScale=0
  Else
    wflash6.intensityscale=1
  end if
  if chase3=0 or swirlramplight3.intensityscale < 0.01 Then
    wflash7.intensityScale=0
  Else
    wflash7.intensityscale=1
  end if
End Sub

'******************************************************
'           TROUGH
'******************************************************

Sub sw58_Hit():Controller.Switch(58) = 1:UpdateTrough:End Sub
Sub sw58_UnHit():Controller.Switch(58) = 0:UpdateTrough:End Sub
Sub sw57_Hit():Controller.Switch(57) = 1:UpdateTrough:End Sub
Sub sw57_UnHit():Controller.Switch(57) = 0:UpdateTrough:End Sub
Sub sw56_Hit():Controller.Switch(56) = 1:UpdateTrough:End Sub
Sub sw56_UnHit():Controller.Switch(56) = 0:UpdateTrough:End Sub

sub reflectionTrigger_Hit
  if holoState=1 then
    ActiveBall.ReflectionEnabled=0
  end if
end sub

sub reflectionTrigger_UnHit
  ActiveBall.ReflectionEnabled=1
end sub

Sub UpdateTrough()
  UpdateTroughTimer.Interval = 300
  UpdateTroughTimer.Enabled = 1
End Sub

Sub UpdateTroughTimer_Timer()
  If sw56.BallCntOver = 0 Then sw57.kick 60, 9
  If sw57.BallCntOver = 0 Then sw58.kick 60, 9
  Me.Enabled = 0
End Sub

'******************************************************
'         DRAIN & RELEASE
'******************************************************

Sub sw55_Hit() 'Drain
  UpdateTrough
  Controller.Switch(55) = 1
  PlaySoundAT "drain", sw55
End Sub

Sub sw55_UnHit()  'Drain
  Controller.Switch(55) = 0
End Sub

Sub SolOuthole(enabled)
  If enabled Then
    sw55.kick 60,20
    PlaySoundAt SoundFX("solenoid",DOFContactors), sw55
  End If
End Sub

Sub ReleaseBall(enabled)
  If enabled Then
    PlaySoundAt SoundFX("ballrelease",DOFContactors), sw56
    sw56.kick 60, 12
    UpdateTrough
  End If
End Sub


'******************************************************
'       RIGHT UPPER KICKER (sw34)
'******************************************************

Sub RightUpperKicker(enabled)
  If enabled Then
    sw34.kickz 0,45,1.56,0
    PlaysoundAt SoundFX("solenoid",DOFContactors), sw34
    If Controller.Switch(34) = True Then
      PlaysoundAt SoundFX("vuk_exit",DOFContactors), sw34
    End If
    Controller.Switch(34) = 0
  End If
End Sub

sub sw34_hit
  PlaySoundAt "kicker_enter_center", sw34
  Controller.Switch(34) = 1
End sub


'******************************************************
'       RIGHT LOWER KICKER (sw37)
'******************************************************

Sub RightLowerKicker(enabled)
  If enabled Then
    sw37.kickz 180,20,1.56, 150
    PlaysoundAt SoundFX("solenoid",DOFContactors), sw37
    If Controller.Switch(37) = True Then
      PlaysoundAt SoundFX("vuk_exit",DOFContactors), sw37
    End If
    Controller.Switch(37) = 0
  End If
End Sub

sub sw37_hit
  PlaySoundAt "kicker_enter_center", sw37
  Controller.Switch(37) = 1
  Stopsound "SubWay"
End sub

Sub sw17scoop_hit()
  PlaySoundAt "scoopenter", sw17scoop
End Sub

Sub sw17_Hit()
  PlaysoundAt  "SubWay", sw17
  vpmTimer.PulseSw 17
End Sub

Sub sw16scoop_hit()
  PlaySoundAt "scoop_enter", sw17scoop
End Sub

Sub sw16_Hit()
  PlaysoundAt  "SubWay", sw16
  vpmTimer.PulseSw 16
End Sub


'******************************************************
'         FLIPPERS
'******************************************************

Sub SolLFlipper(Enabled)
    If Enabled Then
        PlaySound SoundFX("fx_flipperup",DOFFlippers), 0, .67, AudioPan(LeftFlipper), 0.05,0,0,1,AudioFade(LeftFlipper)
        LeftFlipper.RotateToEnd
    Else
        PlaySound SoundFX("fx_flipperdown",DOFFlippers), 0, 1, AudioPan(LeftFlipper), 0.05,0,0,1,AudioFade(LeftFlipper)
        LeftFlipper.RotateToStart
    End If
End Sub
'
Sub SolRFlipper(Enabled)
    If Enabled Then
        PlaySound SoundFX("fx_flipperup",DOFFlippers), 0, .67, AudioPan(RightFlipper), 0.05,0,0,1,AudioFade(RightFlipper)
        RightFlipper.RotateToEnd
    Else
        PlaySound SoundFX("fx_flipperdown",DOFFlippers), 0, 1, AudioPan(RightFlipper), 0.05,0,0,1,AudioFade(RightFlipper)
        RightFlipper.RotateToStart
    End If
End Sub

'******************************************************
'           SWITCHES
'******************************************************

 Sub sw51_Hit():VPMTimer.PulseSw 51:PlaySoundAt "Sensor", sw51: End Sub
 Sub sw52_Hit():VPMTimer.PulseSw 52:PlaySoundAt "Sensor", sw52: End Sub
 Sub sw53_Hit():VPMTimer.PulseSw 53:PlaySoundAt "Sensor", sw53: End Sub
 Sub sw54_Hit():VPMTimer.PulseSw 54:PlaySoundAt "Sensor", sw54: End Sub

 Sub sw15_Hit():controller.switch(15) = 0 :PlaySoundAt "Sensor", sw15: End Sub
 Sub sw15_Unhit():controller.switch(15) = 1 End Sub


 Sub sw25_Hit():VPMTimer.PulseSw 25:PlaySoundAt "Sensor", sw25: End Sub
 Sub sw26_Hit():VPMTimer.PulseSw 26:PlaySoundAt "Sensor", sw26: End Sub
 Sub sw27_Hit():VPMTimer.PulseSw 27:PlaySoundAt "Sensor", sw27: End Sub
 Sub sw28_Hit():VPMTimer.PulseSw 28:PlaySoundAt "Sensor", sw28: End Sub

 Sub sw41_Hit():VPMTimer.PulseSw 41:PlaySoundAt SoundFX("Target",DOFTargets), sw41: End Sub
 Sub sw42_Hit():VPMTimer.PulseSw 42:PlaySoundAt SoundFX("Target",DOFTargets), sw42: End Sub
 Sub sw43_Hit():VPMTimer.PulseSw 43:PlaySoundAt SoundFX("Target",DOFTargets), sw43: End Sub
 Sub sw44_Hit():VPMTimer.PulseSw 44:PlaySoundAt SoundFX("Target",DOFTargets), sw44: End Sub

 Sub sw63_Hit():VPMTimer.PulseSw 63:PlaySoundAt "Gate", sw63: End Sub

 Sub sw35_Hit() 'Right Ramp Enter
  VPMTimer.PulseSw 35
  PlaySoundAt "Gate", sw35
  If ActiveBall.vely < 0 then
    PlaySound "RampEntrywGate", 0, Vol(ActiveBall)*0.5, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  Else
    StopSound "RampEntryGate"
  End If
 End Sub  'Right Ramp Enter

 Sub sw36_Hit()
  VPMTimer.PulseSw 36
  PlaySoundAt "Gate", sw36
  PlaySoundAt "Gate", sw36
  If ActiveBall.vely < 0 then
    PlaySound "RampEntrywGate", 0, Vol(ActiveBall)*0.5, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  Else
    StopSound "RampEntryGate"
  End If
 End Sub  'Left Ramp Enter


 Sub sw61_Hit():VPMTimer.PulseSw 61:End Sub          'Right Ramp Exit
 Sub sw62_Hit():VPMTimer.PulseSw 62:End Sub          'Left Ramp Exit
 Sub sw64_Hit():VPMTimer.PulseSw 64:End Sub          'Upper Ramp Enter
 Sub sw18_Hit():VPMTimer.PulseSw 18:PlaySoundAt "Gate", sw18: End Sub  'Center Shot

 Sub sw65_Hit():VPMTimer.PulseSw 65:End Sub          'Pool switch

 Sub sw66_Hit()   : controller.switch(66) = 1 :PlaySoundAt "Sensor", sw66: End Sub 'Ball Shooter Lane
 Sub sw66_UnHit() : controller.switch(66) = 0 : End Sub


'******************************************************
'         BUMPERS
'******************************************************

Sub Bumper1_Hit()
  PlaySoundAt SoundFX("fx_bumper1",DOFContactors), bumper1
  vpmTimer.PulseSw 45
End Sub

Sub Bumper2_Hit()
  PlaySoundAt SoundFX("fx_bumper1",DOFContactors), bumper2
  vpmTimer.PulseSw 46
End Sub

Sub Bumper3_Hit()
  PlaySoundAt SoundFX("fx_bumper1",DOFContactors), bumper3
  vpmTimer.PulseSw 33
End Sub
'**************************************************************
' Moving Ramp
'**************************************************************

Dim RampAngle, RampDir
RampAngle= 0
RampDir = -1     '1 is up -1 is down dir

SolCallback(23) = "SolRampUp"
SolCallback(26) = "SolRampDown"


Sub SolRampUp(Enabled)
    If Enabled Then
        RampDir = 1
        Controller.Switch(38) = False
        UpdateRamp.Enabled = True
        playsound SoundFX("diverter",DOFContactors)
    End If
End Sub

Sub SolRampDown(Enabled)
    If Enabled Then
        RampDir = -1
        Controller.Switch(38) = true
        UpdateRamp.Enabled = True
        playsound SoundFX("diverter",DOFContactors)
    End If
End sub

Sub UpdateRamp_Timer
    RampAngle = RampAngle + RampDir

    If RampAngle> 0 Then   '110
        RampAngle = 0
    RampDir=-1
        UpdateRamp.Enabled = 0
    collRamp.collidable=false
    End If
    If RampAngle < -25 Then
        RampAngle = -25
    RampDir=1
        UpdateRamp.Enabled = 0
    collRamp.collidable=true
    End If
  moveRamp.rotx=RampAngle
End Sub

Dim RStep, Lstep
Sub RightSlingShot_Slingshot
  vpmTimer.pulsesw 48
    PlaySound SoundFX("right_slingshot",DOFContactors), 0,1, 0.05,0.05 '0,1, AudioPan(RightSlingShot), 0.05,0,0,1,AudioFade(RightSlingShot)
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.TransZ = -20
    RStep = 0
    RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 2:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.TransZ = -10
        Case 3:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.TransZ = 0:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
  vpmTimer.pulsesw 47
    PlaySound SoundFX("left_slingshot",DOFContactors), 0,1, -0.05,0.05 '0,1, AudioPan(LeftSlingShot), 0.05,0,0,1,AudioFade(LeftSlingShot)
    LSling.Visible = 0
    LSling1.Visible = 1
    sling2.TransZ = -20
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 2:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.TransZ = -10
        Case 3:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.TransZ = 0:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub


'******************************************************
'       CREATURE ANIMATION
'******************************************************

SolCallBack(27) = "CreatureMotorMirror"
SolCallBack(28) = "HoloLamp"
'SolCallBack(11) = "CreatureFlash"


dim holoState:holoState=0
dim holoIdx:holoIdx=18
dim posinc:posinc=1
  Sub HoloLamp(enabled)
  If enabled then
    holoState=1
      creature.visible=1
  else
    holoState=0
      creature.visible=0
  end if
  End Sub

   Sub CreatureFlash(enabled)
  If enabled then
    holoState=1
      creature.visible=1
  else
    holoState=0
      creature.visible=0
  end if
 End Sub

 Sub CreatureMotorMirror(enabled)
  if enabled then
    Playsound SoundFX("Motor",DOFShaker),-1
      holoIdx=18
    posinc = 1
    creatureTimer.enabled=1
  else
    Stopsound "Motor"
    holoIdx=18
    posinc = 1
    creatureTimer.enabled=0
  end if
End Sub

 Sub creatureTimer_timer()
  holoIdx=holoIdx+posinc
  creature.visible=holoState
  If holoIdx<1 then holoIdx=1:posinc=1
  If holoIdx>46 then holoIdx=46:posinc=-1
  creature.image = holoIdx
end Sub



sub startUpperLeftWireRamp_Hit
  PlaySoundAtBall "WireRamp"
end Sub

sub endUpperLeftWireRamp_Hit():
  StopSound "WireRamp":
  vpmTimer.AddTimer 150, "PlayBallDrop endUpperLeftWireRamp, 0.1'"
end sub

Sub leftrampdrop_hit()
  vpmTimer.AddTimer 150, "PlayBallDrop leftrampdrop, 0.1'"
End Sub


Sub rightrampdrop_hit()
  vpmTimer.AddTimer 150, "PlayBallDrop rightrampdrop, 0.1'"
End Sub

Sub startCenterWireRamp_Hit
  PlaySoundAtBall "CenterWireRamp"
End Sub

Sub stopCenterWireRamp_Hit
  StopSound "CenterWireRamp"
End Sub


Sub swirlrampdrop_hit()
  vpmTimer.AddTimer 150, "PlayBallDrop swirlrampdrop, 0.1'"
End Sub

Sub PlayBallDrop(obj,volume)
  Playsound "BallDrop", 1, volume, AudioPan(obj), 0,0,0, 1, AudioFade(obj)'"
End Sub

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

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "table1" is the name of the table
    Dim tmp
    tmp = ball.x * 2 / table1.width-1
    If tmp> 0 Then
        Pan = Csng(tmp ^10)
    Else
        Pan = Csng(-((- tmp) ^10) )
    End If
End Function

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
'   rothbauerw's Manual Ball Control
'*****************************************

'************************   Start Ball Control 3/3
Sub StartBallControl_Hit()
  Set ControlBall = ActiveBall
  contballinplay = true
End Sub

Sub StopBallControl_Hit()
  contballinplay = false
End Sub

Dim bcup, bcdown, bcleft, bcright
Dim contball, contballinplay, ControlBall
Dim bcvel, bcyveloffset, bcboostmulti, bcboost

bcboost = 1   'Do Not Change - default setting
bcvel = 4   'Controls the speed of the ball movement
bcyveloffset = -0.01  'Offsets the force of gravity to keep the ball from drifting vertically on the table, should be negative
bcboostmulti = 3  'Boost multiplier to ball veloctiy (toggled with the B key)

Sub BallControlTimer_Timer()
  If Contball and ContBallInPlay and EnableBallControl then
    If bcright = 1 Then
      ControlBall.velx = bcvel*bcboost
    ElseIf bcleft = 1 Then
      ControlBall.velx = - bcvel*bcboost
    Else
      ControlBall.velx=0
    End If

    If bcup = 1 Then
      ControlBall.vely = -bcvel*bcboost
    ElseIf bcdown = 1 Then
      ControlBall.vely = bcvel*bcboost
    Else
      ControlBall.vely= bcyveloffset
    End If
  End If
End Sub
'************************   End Ball Control 3/3

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

Sub BallRollingUpdate_Timer()
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
      StopSound("plasticroll" & b)
            rolling(b) = True
            PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b))*RollingSoundFactor, AudioPan(BOT(b)), 0, Pitch(BOT(b)), 1, 0, AudioFade(BOT(b))
        ElseIf BallVel(BOT(b) ) > 1 AND BOT(b).z > 30 Then
      StopSound("fx_ballrolling" & b)
            rolling(b) = True
            PlaySound("plasticroll" & b), -1, Vol(BOT(b))*RollingSoundFactor*2, AudioPan(BOT(b)), 0, Pitch(BOT(b)), 1, 0, AudioFade(BOT(b))
    Else
            If rolling(b) = True Then
                StopSound("fx_ballrolling" & b)
        StopSound("plasticroll" & b)
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
  batleftshadow.objRotZ = LeftFlipper.currentangle
  batrightshadow.objRotZ = RightFlipper.currentangle
  batleft.objrotz = LeftFlipper.currentangle
  batright.objrotz=RightFlipper.currentangle
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
            BallShadow(b).X = ((BOT(b).X) + ((BOT(b).X - (Table1.Width/2))/40))
        Else
            BallShadow(b).X = ((BOT(b).X) + ((BOT(b).X - (Table1.Width/2))/40))
        End If
        ballShadow(b).Y = BOT(b).Y + 10
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
  PlaySound "pinhit_low", 0, Vol(ActiveBall)*VolumeDial, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Woods_Hit (idx)
  PlaySound "woodhit", 0, Vol(ActiveBall)*VolumeDial*2, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Targets_Hit (idx)
  PlaySound "target", 0, Vol(ActiveBall)*VolumeDial, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Metals_Thin_Hit (idx)
  PlaySound "metalhit_thin", 0, Vol(ActiveBall)*VolumeDial, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Metals_Medium_Hit (idx)
  PlaySound "metalhit_medium", 0, Vol(ActiveBall)*VolumeDial, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Metals2_Hit (idx)
  PlaySound "metalhit2", 0, Vol(ActiveBall)*VolumeDial, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Gates_Hit (idx)
  PlaySound "gate4", 0, Vol(ActiveBall)*VolumeDial, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Spinner_Spin
  PlaySound "fx_spinner", 0, .25*VolumeDial, AudioPan(Spinner), 0.25, 0, 0, 1, AudioFade(Spinner)
End Sub

Sub Rubbers_Hit(idx)
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 20 then
    PlaySound "fx_rubber2", 0, Vol(ActiveBall)*VolumeDial, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End if
  If finalspeed >= 6 AND finalspeed <= 20 then
    RandomSoundRubber()
  End If
End Sub

Sub Posts_Hit(idx)
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 16 then
    PlaySound "fx_rubber2", 0, Vol(ActiveBall)*VolumeDial, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End if
  If finalspeed >= 6 AND finalspeed <= 16 then
    RandomSoundRubber()
  End If
End Sub

Sub RandomSoundRubber()
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySound "rubber_hit_1", 0, Vol(ActiveBall)*VolumeDial, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 2 : PlaySound "rubber_hit_2", 0, Vol(ActiveBall)*VolumeDial, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 3 : PlaySound "rubber_hit_3", 0, Vol(ActiveBall)*VolumeDial, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End Select
End Sub

Sub LeftFlipper_Collide(parm)
  RandomSoundFlipper()
End Sub

Sub RightFlipper_Collide(parm)
  RandomSoundFlipper()
End Sub

Sub DiverterFlipper_Collide(parm)
  PlaySound "metalhit2", 0, Vol(ActiveBall)*VolumeDial, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub RandomSoundFlipper()
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySound "flip_hit_1", 0, Vol(ActiveBall)*VolumeDial, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 2 : PlaySound "flip_hit_2", 0, Vol(ActiveBall)*VolumeDial, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 3 : PlaySound "flip_hit_3", 0, Vol(ActiveBall)*VolumeDial, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End Select
End Sub


'================Light Handling==================
'   GI, Flashers, and Lamp handling
'Based on JP's VP10 fading Lamp routine, based on PD's Fading Lights
'   Mod FrameTime and GI handling by nFozzy
'================================================
'Short installation
'Keep all non-GI lamps/Flashers in a big collection called aLampsAll
'Initialize SolModCallbacks: Const UseVPMModSol = 1 at the top of the script, before LoadVPM. vpmInit me in table1_Init()
'LUT images (optional)
'Make modifications based on era of game (setlamp / flashc for games without solmodcallback, use bonus GI subs for games with only one GI control)

Dim LampState(340), FadingLevel(340), CollapseMe
Dim FlashSpeedUp(340), FlashSpeedDown(340), FlashMin(340), FlashMax(340), FlashLevel(340)
Dim SolModValue(340)  'holds 0-255 modulated solenoid values

'These are used for fading lights and flashers brighter when the GI is darker
Dim LampsOpacity(340, 2) 'Columns: 0 = intensity / opacity, 1 = fadeup, 2 = FadeDown
Dim GIscale(4)  '5 gi strings

InitLamps

reDim CollapseMe(1) 'Setlamps and SolModCallBacks (Click Me to Collapse)
  Sub SetLamp(nr, value)
    If value <> LampState(nr) Then
      LampState(nr) = abs(value)
      FadingLevel(nr) = abs(value) + 4
    End If
  End Sub

  Sub SetLampm(nr, nr2, value)  'set 2 lamps
    If value <> LampState(nr) Then
      LampState(nr) = abs(value)
      FadingLevel(nr) = abs(value) + 4
    End If
    If value <> LampState(nr2) Then
      LampState(nr2) = abs(value)
      FadingLevel(nr2) = abs(value) + 4
    End If
  End Sub

  Sub SetModLamp(nr, value)
    If value <> SolModValue(nr) Then
      SolModValue(nr) = value
      if value > 0 then LampState(nr) = 1 else LampState(nr) = 0
      FadingLevel(nr) = LampState(nr) + 4
    End If
  End Sub

  Sub SetModLampM(nr, nr2, value) 'set 2 modulated lamps
    If value <> SolModValue(nr) Then
      SolModValue(nr) = value
      if value > 0 then LampState(nr) = 1 else LampState(nr) = 0
      FadingLevel(nr) = LampState(nr) + 4
    End If
    If value <> SolModValue(nr2) Then
      SolModValue(nr2) = value
      if value > 0 then LampState(nr2) = 1 else LampState(nr2) = 0
      FadingLevel(nr2) = LampState(nr2) + 4
    End If
  End Sub
  'Flashers via SolModCallBacks
  SolModCallBack(2) = "SetModLamp 102," 'Flasher2
  SolModCallBack(8) = "SetModLamp 108," 'Flasher8
  SolModCallBack(9) = "SetModLamp 109," 'Flasher9
  SolModCallBack(10) = "SetModLamp 110," 'Flasher10
  SolModCallBack(16) = "SetModLamp 116," 'Flasher16
  SolModCallBack(17) = "SetModLamp 117," 'Flasher17
  SolModCallBack(18) = "SetModLamp 118," 'Flasher18
  SolModCallBack(19) = "SetModLamp 119," 'Flasher19
  SolModCallBack(22) = "SetModLamp 122," 'Flasher22
  SolModCallBack(25) = "SetModLamp 125," 'MovieFlasher

'#end section
reDim CollapseMe(2) 'InitLamps  (Click Me to Collapse)
  Sub InitLamps() 'set fading speeds and other stuff here
    GetOpacity aLampsAll  'All non-GI lamps and flashers go in this object array for compensation script!
    Dim x
    for x = 0 to uBound(LampState)
      LampState(x) = 0  ' current light state, independent of the fading level. 0 is off and 1 is on
      FadingLevel(x) = 4  ' used to track the fading state
      FlashSpeedUp(x) = 0.1 'Fading speeds in opacity per MS I think (Not used with nFadeL or nFadeLM subs!)
      FlashSpeedDown(x) = 0.1

      FlashMin(x) = 0.001     ' the minimum value when off, usually 0
      FlashMax(x) = 1       ' the minimum value when off, usually 1
      FlashLevel(x) = 0.001       ' Raw Flasher opacity value. Start this >0 to avoid initial flasher stuttering.

      SolModValue(x) = 0      ' Holds SolModCallback values

    Next

    for x = 0 to uBound(giscale)
      Giscale(x) = 1.625      ' lamp GI compensation multiplier, eg opacity x 1.625 when gi is fully off
    next

    for x = 11 to 100 'insert fading levels (only applicable for lamps that use FlashC sub)
      FlashSpeedUp(x) = 0.015
      FlashSpeedDown(x) = 0.009
    Next

    for x = 101 to 186  'Flasher fading speeds 'intensityscale(%) per 10MS
      FlashSpeedUp(x) = 1.1 * 4
      FlashSpeedDown(x) = 0.9 * 4
    next

    for x = 200 to 204    'GI relay on / off  fading speeds
      FlashSpeedUp(x) = 0.01
      FlashSpeedDown(x) = 0.008
      FlashMin(x) = 0
    Next
    for x = 300 to 304    'GI 8 step modulation fading speeds
      FlashSpeedUp(x) = 0.01
      FlashSpeedDown(x) = 0.008
      FlashMin(x) = 0
    Next

    UpdateGIon 0, 1:UpdateGIon 1, 1: UpdateGIon 2, 1 : UpdateGIon 3, 1:UpdateGIon 4, 1
    UpdateGI 0, 7:UpdateGI 1, 7:UpdateGI 2, 7 : UpdateGI 3, 7:UpdateGI 4, 7
  End Sub

  Sub GetOpacity(a) 'Keep lamp/flasher data in an array
    Dim x
    for x = 0 to (a.Count - 1)
      On Error Resume Next
      if a(x).Opacity > 0 then a(x).Uservalue = a(x).Opacity
      if a(x).Intensity > 0 then a(x).Uservalue = a(x).Intensity
      If a(x).FadeSpeedUp > 0 then LampsOpacity(x, 1) = a(x).FadeSpeedUp : LampsOpacity(x, 2) = a(x).FadeSpeedDown
    Next
    for x = 0 to (a.Count - 1) : LampsOpacity(x, 0) = a(x).UserValue : Next
  End Sub

  sub DebugLampsOn(input):Dim x: for x = 10 to 100 : setlamp x, input : next :  end sub

'#end section

reDim CollapseMe(3) 'LampTimer  (Click Me to Collapse)
  LampTimer.Interval = -1 '-1 is ideal, but it will technically work with any timer interval
  LampTimer.Enabled = 1
  Dim FrameTime, InitFadeTime : FrameTime = 10  'Count Frametime
  Sub LampTimer_Timer()
    FrameTime = gametime - InitFadeTime
    Dim chgLamp, num, chg, ii
    chgLamp = Controller.ChangedLamps
    If Not IsEmpty(chgLamp) Then
      For ii = 0 To UBound(chgLamp)
        LampState(chgLamp(ii, 0) ) = chgLamp(ii, 1)       'keep the real state in an array
        FadingLevel(chgLamp(ii, 0) ) = chgLamp(ii, 1) + 4 'actual fading step
      Next
    End If

    UpdateGIstuff
    UpdateLamps
    UpdateFlashers

    InitFadeTime = gametime
  End Sub
'#end section
reDim CollapseMe(4) 'ASSIGNMENTS: Lamps, GI, and Flashers (Click Me to Collapse)
  Sub UpdateGIstuff()
    FadeGI 200  'Chaselights #1
    ModGI  300
    FadeGI 201  'Middle GI
    ModGI  301
    FadeGI 202  'Upper GI
    ModGI  302
    FadeGI 203  'Chaselights #2
    ModGI  303
    FadeGI 204  'Lower GI
    ModGI  304
    UpdateGIObjects 200, 300, Chaselights, 1  'nr, nr2, array 'Updates GI objects
    UpdateGIObjects 201, 301, GI_Middle, 0  'nr, nr2, array 'Updates GI objects
    UpdateGIObjects 202, 302, GI_Upper, 0 'nr, nr2, array 'Updates GI objects
    UpdateGIObjects 204, 304, GI_Lower, 0 'nr, nr2, array 'Updates GI objects
    If lutfading = 1 Then
      GiCompensationAvgM 201, 301, 202, 302, 204, 304, aLampsAll, GiScale(2)  'Averages two lamp strings. Ideal match the avg. Lut fading.
      FadeLUTavgM 201, 301, 202, 302, 204, 304, "LutCont_", 27  'Lut averages three fading strings
    End If
  End Sub

  Dim LutFading

  If Table1.ShowDT = True Then
    lutfading = dtlutfading
  Else
    lutFading = fslutfading
  End If

  Sub UpdateFlashers()
    nModFlash 102, fl2, 0, 0  'NR, Object, LowPass(see function ScaleByte) OffScale (multiply fading when callback 0)
    nModFlash 108, fl8, 0, 0  'Last two arguments will carry through to objects controlled by nModFlashM.
    nModFlashm 109, fl9b
    nModFlash 109, fl9a, 0, 0 'For more control, use multiple fading numbers via SetModLampM / SetModLampMM
    nModFlashm 110, fl10a
    nModFlash 110, fl10, 0, 0
    nModFlash 116, fl16, 0, 0
    nModFlash 117, fl17, 0, 0
    nModFlash 118, fl18, 0, 0
    nModFlash 119, fl19, 0, 0
    nModFlash 122, fl22, 0, 0
    nModFlashm 125, fl37b
    nModFlashm 125, fl37a
    nModFlashm 125, l44b
    nModFlash 125, l37b, 0, 0
  End Sub

  Sub UpdateLamps
    NFadeLm 11,  l11
    NFadeLm 11,  light51
    FlashC   11,  wFlash11
    NFadeLm 12,  l12
    NFadeLm 12,  light52
    FlashC   12,  wFlash12
    NFadeLm 13,  l13
    NFadeLm 13,  light53
    FlashC   13,  wFlash13
    NFadeLm 14,  l14
    NFadeLm 14,  light54
    FlashC   14,  wFlash14

'flupper
    If bumperslit = 0 then
      NFadeLm 15,  l15
      NFadeLm 15,  light2
      NFadeLm 16,  l16
      NFadeLm 16,  light1
      NFadeLm 17,  l17
      NFadeLm 17,  light4
    Else
      NFadeLm 15, bumper2Light1
      NFadeLm 15, bumper2Light2
      NFadeLm 16, bumper1Light1
      NFadeLm 16, bumper1Light2
      NFadeLm 17, bumper3Light1
      NFadeLm 17, bumper3Light2
    end if

'flupper
    FlashC 18,  f18
    NFadeL 21,  l21
    NFadeL 22,  l22
    NFadeL 23,  l23
    NFadeL 24,  l24
    NFadeLm 25,  l25
    NFadeL 25,  light34
    NFadeLm 26,  l26
    NFadeL 26,  light35
    NFadeL 27,  l27
    NFadeLm 28,  l28
    NFadeLm 28,  l288
    NFadeL 28,  light38
    NFadeLm 31,  l31
    NFadeL 31,  light50
    NFadeL 32,  l32
    NFadeL 33,  l33
    NFadeLm 34,  l34
    NFadeL 34,  light45
    NFadeLm 35,  l35
    NFadeLm 35,  light44
    FlashC  35,  wflash44
    NFadeLm 36,  l36
    NFadeLm 36,  light42
    NFadeLm 36,  light47
    flashC  36,  wflash42
    NFadeLm 37,  l37a
'   NFadeLm 37,  l37b
    NFadeLm 37,  l37c
    NFadeLm 37,  light39
    NFadeLm 37,  light40
    FlashC  40,  wflash40
    NFadeLm 37,  light41
    NFadeLm 37,  light46
    FlashC  37,  wflash41
    NFadeLm 38,  l38
    NFadeL 38,  light27
    NFadeLm 41,  l41
    NFadeL 41,  light33
    NFadeLm 42,  l42
    NFadeLm 42,  light32
    FlashC   42,  wflash32
    NFadeLm 43,  l43
    NFadeL 43,  light31
    NFadeLm 44,  l44a
'   NFadeLm 44,  l44b
    NFadeLm 44,  l44c
    NFadeLm 44,  light28
    NFadeLm 44,  light29
    NFadeLm 44,  light30
    NFadeLm 44,  light55
    FlashC   44,  wflash30
    NFadeL 45,  l45
    NFadeLm 46,  l46
    NFadeL 46,  light43
    NFadeLm 47,  l47
    NFadeL 47,  light22
    NFadeLm 48,  l48
    NFadeL 48,  light21
    NFadeLm 54,  l54
    NFadeL 54,  light23
    NFadeLm 53,  l53
    NFadeL 53,  light24
    NFadeL 52,  l52
    FlashC 51,  f51

    FlashC 55, fl55
    FlashC 56, fl56
    FlashC 57, fl57
    FlashC 58, fl58

    NFadeL 61,  l61
    NFadeLm 62,  l62
    NFadeL 62,  light36
    Flashm 63, wFlash63
    FlashC 63, fl63
    Flashm 64, wFlash64
    FlashC 64, fl64
    Flashm 65, wFlash65
    FlashC 65, fl65
    FlashC 66, fl66
    FlashC 67, fl67
    FlashC 68, fl68

    NFadeL 81,  l81
    NFadeL 82,  l82
    NFadeL 83,  l83
    NFadeL 84,  l84
    NFadeLm 85,  l85
    NFadeLm 85,  light37
    flashC  85,  wflash37
    NFadeLm 86,  l86
    NFadeLm 86,  light26
    NFadeL 86,  light49
    NFadeLm 87,  l87
    NFadeLm 87,  light25
    NFadeL 87,  light48
  end sub

'#end section

reDim CollapseMe(5) 'Combined GI subs / functions (Click Me to Collapse)
  Set GICallback = GetRef("UpdateGIon")   'On/Off GI to NRs 200-203
  Sub UpdateGIOn(no, Enabled) : Setlamp no+200, cInt(enabled) : End Sub


  Set GICallback2 = GetRef("UpdateGI")
  Sub UpdateGI(no, step)            '8 step Modulated GI to NRs 300-303
    Dim ii, x', i
    If step = 0 then exit sub 'only values from 1 to 8 are visible and reliable. 0 is not reliable and 7 & 8 are the same so...

    If no = 0 Then
      If step > 0 then chaselightson = true else chaselightson = false : end if
    end if

    SetModLamp no+300, ScaleGI(step, 0)
    LampState((no+300)) = 0
  End Sub

  Function ScaleGI(value, scaletype)  'returns an intensityscale-friendly 0->100% value out of 1>8 'it does go to 8
    Dim i
    Select Case scaletype 'select case because bad at maths
      case 0  : i = value * (1/8) '0 to 1
      case 25 : i = (1/28)*(3*value + 4)
      case 50 : i = (value+5)/12
      case else : i = value * (1/8) '0 to 1
  '     x = (4*value)/3 - 85  '63.75 to 255
    End Select
    ScaleGI = i
  End Function

' Dim LSstate : LSstate = False 'fading sub handles SFX 'Uncomment to enable
  Sub FadeGI(nr) 'in On/off   'Updates nothing but flashlevel
    Select Case FadingLevel(nr)
      Case 3 : FadingLevel(nr) = 0
      Case 4 'off
        DOF 101, DOFOff
  '     If Not LSstate then Playsound "FX_Relay_Off",0,LVL(0.1) : LSstate = True  'handle SFX
        FlashLevel(nr) = FlashLevel(nr) - (FlashSpeedDown(nr) * FrameTime)
        If FlashLevel(nr) < FlashMin(nr) Then
           FlashLevel(nr) = FlashMin(nr)
           FadingLevel(nr) = 3 'completely off
  '       LSstate = False
        End if
      Case 5 ' on
        DOF 101, DOFOn
  '     If Not LSstate then Playsound "FX_Relay_On",0,LVL(0.1) : LSstate = True 'handle SFX
        FlashLevel(nr) = FlashLevel(nr) + (FlashSpeedUp(nr) * FrameTime)
        If FlashLevel(nr) > FlashMax(nr) Then
          FlashLevel(nr) = FlashMax(nr)
          FadingLevel(nr) = 6 'completely on
  '       LSstate = False
        End if
      Case 6 : FadingLevel(nr) = 1
    End Select
  End Sub
  Sub ModGI(nr2) 'in 0->1   'Updates nothing but flashlevel 'never off
    Dim DesiredFading
    Select Case FadingLevel(nr2)
      case 3 : FadingLevel(nr2) = 0 'workaround - wait a frame to let M sub finish fading
  '   Case 4 : FadingLevel(nr2) = 3 'off -disabled off, only gicallback1 can turn off GI(?) 'experimental
      Case 5, 4 ' Fade (Dynamic)
        DesiredFading = SolModValue(nr2)
        if FlashLevel(nr2) < DesiredFading Then '+
          FlashLevel(nr2) = FlashLevel(nr2) + (FlashSpeedUp(nr2)  * FrameTime )
          If FlashLevel(nr2) >= DesiredFading Then FlashLevel(nr2) = DesiredFading : FadingLevel(nr2) = 1
        elseif FlashLevel(nr2) > DesiredFading Then '-
          FlashLevel(nr2) = FlashLevel(nr2) - (FlashSpeedDown(nr2) * FrameTime  )
          If FlashLevel(nr2) <= DesiredFading Then FlashLevel(nr2) = DesiredFading : FadingLevel(nr2) = 6
        End If
      Case 6
        FadingLevel(nr2) = 1
    End Select
  End Sub

  Sub UpdateGIobjects(nr, nr2, a, var)  'Just Update GI
    If FadingLevel(nr) > 1 or FadingLevel(nr2) > 1 Then
      Dim x, Output : Output = FlashLevel(nr2) * FlashLevel(nr)
      If var = 1 then Output = Output*chaselightsintensity
      for each x in a : x.IntensityScale = Output : next
    End If
  end Sub

  Sub GiCompensation(nr, nr2, a, GIscaleOff)  'One NR pairing only fading
  ' tbgi.text = "GI: " & SolModValue(nr) & " " & FlashLevel(nr) & " " & FadingLevel(nr) & vbnewline & _
  '       "ModGI: " & SolModValue(nr2) & " " & FlashLevel(nr2) & " " & FadingLevel(nr2) & vbnewline & _
  '       "Solmodvalue, Flashlevel, Fading step"
    if FadingLevel(nr) > 1 or FadingLevel(nr2) > 1 Then
      Dim x, Giscaler, Output : Output = FlashLevel(nr2) * FlashLevel(nr)
      Giscaler = ((Giscaleoff-1) * (ABS(Output-1) )  ) + 1  'fade GIscale the opposite direction

      for x = 0 to (a.Count - 1) 'Handle Compensate Flashers
        On Error Resume Next
        a(x).Opacity = LampsOpacity(x, 0) * Giscaler
        a(x).Intensity = LampsOpacity(x, 0) * Giscaler
        a(x).FadeSpeedUp = LampsOpacity(x, 1) * Giscaler
        a(x).FadeSpeedDown = LampsOpacity(x, 2) * Giscaler
      Next
      '   tbbb.text = giscaler & " on:" & FadingLevel(nr) & vbnewline & "flash: " & output & " onmod:" & FadingLevel(nr2) & vbnewline & l37.intensity
      '   tbbb1.text = FadingLevel(nr) & vbnewline & FadingLevel(nr2)
      ' tbgi1.text = Output & " giscale:" & giscaler  'debug
    End If
    '   tbbb1.text = FLashLevel(nr) & vbnewline & FlashLevel(nr2)
  End Sub

  Sub GiCompensationAvg(nr, nr2, nr3, nr4, a, GIscaleOff) 'Two pairs of NRs averaged together
  ' tbgi.text = "GI: " & SolModValue(nr) & " " & FlashLevel(nr) & " " & FadingLevel(nr) & vbnewline & _
  '       "ModGI: " & SolModValue(nr2) & " " & FlashLevel(nr2) & " " & FadingLevel(nr2) & vbnewline & _
  '       "Solmodvalue, Flashlevel, Fading step"
    if FadingLevel(nr) > 1 or FadingLevel(nr2) > 1 or FadingLevel(nr3) > 1 or FadingLevel(nr4) > 1 Then
      Dim x, Giscaler, Output : Output = (((FlashLevel(nr)*FlashLevel(nr2)) + (FlashLevel(nr3)*Flashlevel(nr4))) /2)
      Giscaler = ((Giscaleoff-1) * (ABS(Output-1) )  ) + 1  'fade GIscale the opposite direction

      for x = 0 to (a.Count - 1) 'Handle Compensate Flashers
        On Error Resume Next
        a(x).Opacity = LampsOpacity(x, 0) * Giscaler
        a(x).Intensity = LampsOpacity(x, 0) * Giscaler
        a(x).FadeSpeedUp = LampsOpacity(x, 1) * Giscaler
        a(x).FadeSpeedDown = LampsOpacity(x, 2) * Giscaler
      Next


    REM tbgi1.text = "Output:" & output & vbnewline & _
          REM "GIscaler" & giscaler & vbnewline & _
          REM "..."
    End If
    REM tbgi.text = "GI0 " & flashlevel(200) & " " & flashlevel(300) & vbnewline & _
          REM "GI1 " & flashlevel(201) & " " & flashlevel(301) & vbnewline & _
          REM "GI2 " & flashlevel(202) & " " & flashlevel(302) & vbnewline & _
          REM "GI3 " & flashlevel(203) & " " & flashlevel(303) & vbnewline & _
          REM "GI4 " & flashlevel(204) & " " & flashlevel(304) & vbnewline & _
          REM "..."
  End Sub

  Sub GiCompensationAvgM(nr, nr2, nr3, nr4, nr5, nr6, a, GIscaleOff)  'Three pairs of NRs averaged together
  ' tbgi.text = "GI: " & SolModValue(nr) & " " & FlashLevel(nr) & " " & FadingLevel(nr) & vbnewline & _
  '       "ModGI: " & SolModValue(nr2) & " " & FlashLevel(nr2) & " " & FadingLevel(nr2) & vbnewline & _
  '       "Solmodvalue, Flashlevel, Fading step"
    if FadingLevel(nr) > 1 or FadingLevel(nr2) > 1 Then
      Dim x, Giscaler, Output
      Output = (((FlashLevel(nr)*FlashLevel(nr2)) + (FlashLevel(nr3)*Flashlevel(nr4)) + (FlashLevel(nr5)*FlashLevel(nr6)))/3)

      Giscaler = ((Giscaleoff-1) * (ABS(Output-1) )  ) + 1  'fade GIscale the opposite direction

      for x = 0 to (a.Count - 1) 'Handle Compensate Flashers
        On Error Resume Next
        a(x).Opacity = LampsOpacity(x, 0) * Giscaler
        a(x).Intensity = LampsOpacity(x, 0) * Giscaler
        a(x).FadeSpeedUp = LampsOpacity(x, 1) * Giscaler
        a(x).FadeSpeedDown = LampsOpacity(x, 2) * Giscaler
      Next
      '   tbbb.text = giscaler & " on:" & FadingLevel(nr) & vbnewline & "flash: " & output & " onmod:" & FadingLevel(nr2) & vbnewline & l37.intensity
      '   tbbb1.text = FadingLevel(nr) & vbnewline & FadingLevel(nr2)
      ' tbgi1.text = Output & " giscale:" & giscaler  'debug
    End If
    '   tbbb1.text = FLashLevel(nr) & vbnewline & FlashLevel(nr2)
  End Sub

  Sub FadeLUT(nr, nr2, LutName, LutCount) 'fade lookuptable NOTE- this is a bad idea for darkening your table as
    If FadingLevel(nr) >2 or FadingLevel(nr2) > 2 Then        '-it will strip the whites out of your image
      Dim GoLut
      GoLut = cInt(LutCount * (FlashLevel(nr)*FlashLevel(nr2) ) )
      Table1.ColorGradeImage = LutName & GoLut
  '   tbgi2.text = Table1.ColorGradeImage & vbnewline & golut 'debug
    End If
  End Sub

  Sub FadeLUTavg(nr, nr2, nr3, nr4, LutName, LutCount)  'FadeLut for two GI strings (WPC)
    If FadingLevel(nr) >2 or FadingLevel(nr2) > 2 or FadingLevel(nr3) >2 or FadingLevel(nr4) > 2 Then
      Dim GoLut
      GoLut = cInt(LutCount * (((FlashLevel(nr)*FlashLevel(nr2)) + (FlashLevel(nr3)*Flashlevel(nr4))) /2) )
      Table1.ColorGradeImage = LutName & GoLut
      REM tbgi2.text = Table1.ColorGradeImage & vbnewline & golut 'debug
    End If
  End Sub

  Sub FadeLUTavgM(nr, nr2, nr3, nr4, nr5, nr6, LutName, LutCount) 'FadeLut for three GI strings (WPC)
    If FadingLevel(nr) >2 or FadingLevel(nr2) > 2 or FadingLevel(nr3) >2 or FadingLevel(nr4) > 2 or _
    FadingLevel(nr5) >2 or FadingLevel(nr6) > 2 Then
      Dim GoLut
      GoLut = cInt(LutCount * (((FlashLevel(nr)*FlashLevel(nr2)) + (FlashLevel(nr3)*Flashlevel(nr4)) + (FlashLevel(nr5)*FlashLevel(nr6)))/3)  ) 'what a mess
      Table1.ColorGradeImage = LutName & GoLut
  '   tbgi2.text = Table1.ColorGradeImage & vbnewline & golut 'debug
    End If
  End Sub

'#end section

reDim CollapseMe(6) 'Fading subs   (Click Me to Collapse)
  Sub nModFlash(nr, object, scaletype, offscale)  'Fading with modulated callbacks
    Dim DesiredFading
    Select Case FadingLevel(nr)
      case 3 : FadingLevel(nr) = 0  'workaround - wait a frame to let M sub finish fading
      Case 4  'off
        If Offscale = 0 then Offscale = 1
        FlashLevel(nr) = FlashLevel(nr) - (FlashSpeedDown(nr) * FrameTime ) * offscale
        If FlashLevel(nr) < 0 then FlashLevel(nr) = 0 : FadingLevel(nr) = 3
        Object.IntensityScale = ScaleLights(FlashLevel(nr),0 )
      Case 5 ' Fade (Dynamic)
        DesiredFading = ScaleByte(SolModValue(nr), scaletype)
        if FlashLevel(nr) < DesiredFading Then '+
          FlashLevel(nr) = FlashLevel(nr) + (FlashSpeedUp(nr) * FrameTime )
          If FlashLevel(nr) >= DesiredFading Then FlashLevel(nr) = DesiredFading : FadingLevel(nr) = 1
        elseif FlashLevel(nr) > DesiredFading Then '-
          FlashLevel(nr) = FlashLevel(nr) - (FlashSpeedDown(nr) * FrameTime )
          If FlashLevel(nr) <= DesiredFading Then FlashLevel(nr) = DesiredFading : FadingLevel(nr) = 6
        End If
        Object.Intensityscale = ScaleLights(FlashLevel(nr),0 )
      Case 6 : FadingLevel(nr) = 1
    End Select
  End Sub

  Sub nModFlashM(nr, Object)
    Select Case FadingLevel(nr)
      Case 3, 4, 5, 6 : Object.Intensityscale = ScaleLights(FlashLevel(nr),0 )
    End Select
  End Sub

  Sub Flashc(nr, object)  'FrameTime Compensated. Can work with Light Objects (make sure state is 1 though)
    Select Case FadingLevel(nr)
      Case 3 : FadingLevel(nr) = 0
      Case 4 'off
        FlashLevel(nr) = FlashLevel(nr) - (FlashSpeedDown(nr) * FrameTime)
        If FlashLevel(nr) < FlashMin(nr) Then
          FlashLevel(nr) = FlashMin(nr)
           FadingLevel(nr) = 3 'completely off
        End if
        Object.IntensityScale = FlashLevel(nr)
      Case 5 ' on
        FlashLevel(nr) = FlashLevel(nr) + (FlashSpeedUp(nr) * FrameTime)
        If FlashLevel(nr) > FlashMax(nr) Then
          FlashLevel(nr) = FlashMax(nr)
          FadingLevel(nr) = 6 'completely on
        End if
        Object.IntensityScale = FlashLevel(nr)
      Case 6 : FadingLevel(nr) = 1
    End Select
  End Sub

  Sub Flashm(nr, object) 'multiple flashers, it just sets the flashlevel
    select case FadingLevel(nr)
      case 3, 4, 5, 6 : Object.IntensityScale = FlashLevel(nr)
    end select
  End Sub

  Sub NFadeL(nr, object)  'Simple VPX light fading using State
    Select Case FadingLevel(nr)
        Case 3:object.state = 0:FadingLevel(nr) = 0
        Case 4:object.state = 0:FadingLevel(nr) = 3
        Case 5:object.state = 1:FadingLevel(nr) = 6
        Case 6:object.state = 1:FadingLevel(nr) = 1
    End Select
  End Sub

  Sub NFadeLm(nr, object) ' used for multiple lights
    Select Case FadingLevel(nr)
      Case 3:object.state = 0
      Case 4:object.state = 0
      Case 5:object.state = 1
      Case 6:object.state = 1
    End Select
  End Sub

'#End Section

reDim CollapseMe(7) 'Fading Functions (Click Me to Collapse)
  Function ScaleLights(value, scaletype)  'returns an intensityscale-friendly 0->100% value out of 255
    Dim i
    Select Case scaletype 'select case because bad at maths   'TODO: Simplify these functions. B/c this is absurdly bad.
      case 0  : i = value * (1 / 255) '0 to 1
      case 6  : i = (value + 17)/272  '0.0625 to 1
      case 9  : i = (value + 25)/280  '0.089 to 1
      case 15 : i = (value / 300) + 0.15
      case 20 : i = (4 * value)/1275 + (1/5)
      case 25 : i = (value + 85) / 340
      case 37 : i = (value+153) / 408   '0.375 to 1
      case 40 : i = (value + 170) / 425
      case 50 : i = (value + 255) / 510 '0.5 to 1
      case 75 : i = (value + 765) / 1020  '0.75 to 1
      case Else : i = 10
    End Select
    ScaleLights = i
  End Function

  Function ScaleByte(value, scaletype)  'returns a number between 1 and 255
    Dim i
    Select Case scaletype
      case 0 : i = value * 1  '0 to 1
      case 9 : i = (5*(200*value + 1887))/1037 'ugh
      case 15 : i = (16*value)/17 + 15
      Case 63 : i = (3*(value + 85))/4
      case else : i = value * 1 '0 to 1
    End Select
    ScaleByte = i
  End Function

Sub theend() : End Sub



REM Troubleshooting :
REM Flashers/gi are intermittent or aren't showing up
REM Ensure flashers start visible, light objects start with state = 1

REM No lamps or no GI
REM Make sure these constants are set up this way
REM Const UseSolenoids = 1
REM Const UseLamps = 0
REM Const UseGI = 1

REM SolModCallback error
REM Ensure you have the latest scripts. Clear out any loose scripts in your tables that might be causing conflicts.

REM Table1 Error
REM Rename the table to Table1 or find/Replace table1 with whatever the table's name is

REM SolModCallbacks aren't sending anything
REM Two important things to get SolModCallbacks to initialize properly:
REM Put this at the top of the script, before LoadVPM
REM Const UseVPMModSol = 1
REM Put this in the table1_Init() section
REM vpmInit me

sub trigger1_hit()
  if activeball.vely < -18 then
    set shooterball = activeball
    trigger1.timerenabled = true
  end if
End Sub

Dim counter, shooterball

Sub Trigger1_timer()
  counter = counter + 1
  if counter > 10 and shooterball.velx < -15 and shooterball.x > 400 and shooterball.x < 740 then
    shooterball.y = 69
  end if

  if counter = 20 Then
    counter = 0
    trigger1.timerenabled = False
  end If
End Sub

Option Explicit
Randomize

' Thalamus 2018-07-24
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Changed UseSolenoids=1 to 2
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

Dim LightHalo_ON, Apron_Insert_Color, GI_Color, NightMod, Flipper_Color, ClearPlastics, CaptiveLightL, CaptiveLightR, RomSet, cGameName, BallRadius, BallMass, ROLCheat
Dim DesktopMode: DesktopMode = MGOW.ShowDT

'*************************************************
'******************* Options *********************
'*************************************************

Apron_Insert_Color = 2 '1 black   2 Brown

GI_Color = 1 '1 = white  2 = red

Flipper_Color = 1 '1 = red 2 = black

'Set Captive Hole light Colors
CaptiveLightL = 1 '0=off 1=Blue 2=Orange 3=Green 4=Purple 5=Red 6=White 7=Yellow
CaptiveLightR = 1 '0=off 1=Blue 2=Orange 3=Green 4=Purple 5=Red 6=White 7=Yellow

ClearPlastics = 0  ' set to 1 to make plastics see though

'Ball Size and Weight
BallRadius = 26
BallMass = 1

'Cheaters
ROLCheat = 0 'Adds post to right out lane

'*****************************
'Rom Version Selector
'*****************************
'1 "mars"   ' 6-digit
'2 "mars7"   ' 7-digit
RomSet = 1

'*************************************************
'*************** End of Options ******************
'*************************************************


' Thing not working/not needed  Please ignore :P

NightMod = 1
LightHalo_ON = 1 'Set to 0 to turn off light halos

'****************************************
'Check the selected ROM version
'****************************************
If RomSet = 1 then cGameName="mars":DisplayTimer.Enabled = true End If
If RomSet = 2 then cGameName="mars7":DisplayTimer7.Enabled = true End If

dim DTbankC,DTbankR
dim x,y,i,switch,diff,RampUp,aux1
Dim MaxBalls, InitTime, EjectTime, TroughEject, TroughCount, iBall, fgBall, BallsInPlay

Const SFlipperOn="FlipperUp"
Const SFlipperOff="FlipperDown"
Const sCoin="coin"
Const UseSolenoids  = 2
Const UseLamps    = False
Const UseGI     = False
const sLSaucer    = 1
const sRSaucer    = 2
const sCDrops   = 5
const sRDrops   = 6
const sKnocker    = 8
const sOutHole    = 9

LoadVPM "01001100","sys80.vbs",2.33

Set LampCallback=GetRef("UpdateLamps")

SolCallback(sKnocker) = "vpmSolSound SoundFX(""Knocker"",DOFKnocker),"
SolCallback(sOutHole) = "ThruLoad"
SolCallback(sCDrops)  = "ResetCDrops"
SolCallback(sRDrops)  = "ResetRDrops"
SolCallback(sLSaucer) = "TSKickit"
SolCallback(sRSaucer) = "RSKickit"
SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
     If Enabled Then
         PlaySoundAtVol SoundFx("mgow flipper up left",DOFFlippers), LeftFlipper, VolFlip:LeftFlipper.RotateToEnd:LeftFlipper2.RotateToEnd
     Else
         PlaySoundAtVol SoundFx("mgow flipper down left",DOFFlippers), LeftFlipper, VolFlip:LeftFlipper.RotateToStart:LeftFlipper2.RotateToStart
     End If
  End Sub

  Sub SolRFlipper(Enabled)
     If Enabled Then
         PlaySoundAtVol SoundFx("mgow flipper up right",DOFFlippers), RightFlipper, VolFlip:RightFlipper.RotateToEnd:RightFlipper2.RotateToEnd
     Else
         PlaySoundAtVol SoundFx("mgow flipper down right",DOFFlippers), RightFlipper, VolFlip:RightFlipper.RotateToStart:RightFlipper2.RotateToStart
     End If
  End Sub

Sub MGOW_Init()
   With Controller
      .GameName=cGameName
      .SplashInfoLine="Mars God of War" & vbNewLine & "Gottlieb 1981"
      .HandleKeyboard=true
      .ShowTitle=false
      .ShowDMDOnly=true
    If DesktopMode = true then .hidden = true Else .hidden = false End If
      .ShowFrame=false
      .HandleKeyboard=False
      .Run
      On Error Resume Next
      If Err Then MsgBox Err.Description
      On Error Goto 0
   End With

   vpmnudge.tiltobj=array(bumper1,bumper2,bumper3,bumper4,sling1)
   vpmnudge.tiltswitch=57
   vpmnudge.sensitivity=5
   OFFLight()
   Ramp5.Collidable=0

  StartLampTimer
  SetGIColor
  CheckDayNight
  SetInstructions
  SetFlipperColors
  SetCaptiveLightColor
  SetClearPlastics
  SetCheats

  Backdrop_Init
' ball through system
  MaxBalls=3
  InitTime=61
  EjectTime=0
  TroughEject=1
  TroughCount=0
  iBall = 3
  fgBall = false
  CheckMaxBalls
    CreatBalls

'''''''''''Reset Slings

' Sling1posta.IsDropped = False
' Sling1postb.IsDropped = True
' Sling1postc.IsDropped = True
' Sling1postd.IsDropped = True
End Sub

Dim sw63active

Sub UpdateMultipleLamps

' Ramp
  If controller.lamp(8)=true then
    RampUp=true
    Ramp5.Collidable=0
    Ramp5a.Collidable=1
    TubeFlasher.enabled = 0
    TubeFlasherOff
    rampDir = 1
    If switch=true then
    SlowLight()
    switch=false
    End If
    rampMoveTimer.Enabled = 1
  Else
    RampUp=false
    Ramp5.Collidable=1
    Ramp5a.Collidable=0
    TubeFlasher.enabled = 1
    rampDir = -1
    If switch=false then
    FastLight()
    switch=true
    End If
    rampMoveTimer.Enabled = 1
  End If

  If controller.lamp(12)=true then
    If sw63active = 0 then
      SW63.timerenabled = 1
      sw63active = 1
    End If
  End If
  If controller.lamp(13)=true then
    KickBallToLane
  End If
End Sub

''''''''Move Ramp
Const rampMoveMax = 16
Dim rampDir
Sub rampMoveTimer_Timer
  Dim tempa
  tempa = ramp.objrotx + 1 * rampDir
  if tempa > rampMoveMax then
    tempa = rampMoveMax
    rampMoveTimer.Enabled = 0
  elseif tempa < 0 then
    tempa = 0
    rampMoveTimer.Enabled = 0
  end if
  ramp.ObjRotX = tempa:pRampLifter.RotX = tempa*-5 + 150:'pRampLifter.TransX = tempa*-.25:
End Sub

'War Bases

Sub TSKickit(enabled)
  If enabled Then
  SW62.Kick 80,5 + Diff
  SW62.timerenabled = true
  Controller.Switch(62)=0
  PlaySoundAtVol SoundFX("mgow kicker out",DOFContactors), SW62, VolKick
  SW60.timerenabled = true
  End If
End Sub

Sub RSKickit(enabled)
  If enabled Then
  SW61.Kick 180,5 + Diff
  SW61.timerenabled = true
  Controller.Switch(61)=0
  PlaySoundAtVol SoundFX("mgow kicker out",DOFContactors), SW61, VolKick

  End If
End Sub

'Through
Sub ThruLoad(enabled)
  If enabled Then
  Drain.Kick 70,15
  Controller.Switch(64)=0
  PlaySoundAtVol SoundFX("Feedtru",DOFContactors), Kicker1, 1
  End If
End Sub

'******************************** SWITCH HANDLING *************************
Sub SW12_Hit():Controller.Switch(12)=1:End Sub
Sub SW12_UnHit():Controller.Switch(12)=0:End Sub
Sub SW44_Hit():Controller.Switch(44)=1:End Sub
Sub SW44_UnHit():Controller.Switch(44)=0:End Sub

'warebase kickers
Sub SW61_Hit():Controller.Switch(61)=1:GetDiff():playsoundAtVol "mgow kicker in", ActiveBall, 1:End Sub
Sub SW62_Hit():Controller.Switch(62)=1:GetDiff():playsoundAtVol "mgow kicker in", ActiveBall, 1:End Sub

'launch kicker
Sub SW63_Hit():pKickRod.ObjRotX = 15:Controller.Switch(63)=1:GetDiff():playsoundAtVol "mgow kicker in", ActiveBall, 1:End Sub
Sub SW64_Hit():Controller.Switch(64)=1:PlaySoundAtVol "Outhole", ActiveBall, 1:End Sub

'scoring slings

Sub SW73a_Slingshot():vpmtimer.pulsesw 73:End Sub
Sub SW73b_Slingshot():vpmtimer.pulsesw 73:End Sub
Sub SW73c_Slingshot():vpmtimer.pulsesw 73:Rubber22a.visible = false:LeafSwitch1b.ObjRotX = -3:Rubber22b.visible = true:me.timerenabled = 1:End Sub
Sub SW73d_Slingshot():vpmtimer.pulsesw 73:End Sub
Sub SW73e_Slingshot():vpmtimer.pulsesw 73:Rubber8a.visible = false:LeafSwitch2b.ObjRotX = -3:Rubber8b.visible = true:me.timerenabled = 1:End Sub
Sub SW73f_Slingshot():vpmtimer.pulsesw 73:End Sub
'Sub SW73g_Slingshot():vpmtimer.pulsesw 73:End Sub
Sub SW73h_Slingshot():vpmtimer.pulsesw 73:Rubber25a.visible = false:LeafSwitch3b.ObjRotX = -3:Rubber25b.visible = true:me.timerenabled = 1:End Sub
Sub SW73i_Slingshot():vpmtimer.pulsesw 73:End Sub
Sub SW73j_Slingshot():vpmtimer.pulsesw 73:End Sub
Sub SW73k_Slingshot():vpmtimer.pulsesw 73:End Sub
Sub SW73l_Slingshot():vpmtimer.pulsesw 73:Rubber22a.visible = false:LeafSwitch5b.ObjRotX = -3:Rubber22d.visible = true:me.timerenabled = 1:End Sub
Sub SW73m_Slingshot():vpmtimer.pulsesw 73:Rubber22a.visible = false:LeafSwitch4b.ObjRotX = -3:Rubber22c.visible = true:me.timerenabled = 1:End Sub
Sub SW73n_Slingshot():vpmtimer.pulsesw 73:End Sub

Sub SW73c_timer:Rubber22b.visible = false:LeafSwitch1b.ObjRotX = 0:Rubber22a.visible = true:me.timerenabled = 0 End Sub
Sub SW73e_timer:Rubber8b.visible = false:LeafSwitch2b.ObjRotX = 0:Rubber8a.visible = true:me.timerenabled = 0 End Sub
Sub SW73h_timer:Rubber25b.visible = false:LeafSwitch3b.ObjRotX = 0:Rubber25a.visible = true:me.timerenabled = 0 End Sub

Sub SW73l_timer:Rubber22d.visible = false:LeafSwitch5b.ObjRotX = 0:Rubber22a.visible = true:me.timerenabled = 0 End Sub
Sub SW73m_timer:Rubber22c.visible = false:LeafSwitch4b.ObjRotX = 0:Rubber22a.visible = true:me.timerenabled = 0 End Sub

Sub SlowLight()
  PlaySound SoundFX("PostUp",DOFGear)
  TubeFlasher.Interval=500
End Sub

Sub FastLight()
  PlaySound SoundFX("RampDown",DOFGear)
  TubeFlasher.Interval=250
End Sub

Sub OFFLight()
End Sub

'---------------------------------------- KEYS

Sub MGOW_KeyDown(ByVal keycode)
  If keycode = PlungerKey Then
    Plunger.PullBack
  End If
  If keycode = RightFlipperKey Then
  Controller.Switch(34)=1
  End If
    If vpmKeyDown(keycode) then Exit Sub
End Sub

Sub MGOW_KeyUp(ByVal keycode)
  If keycode = PlungerKey Then
    Plunger.Fire
    PlaySoundAtVol "Plunger", Plunger, 1
  End If
  If keycode = RightFlipperKey Then
  Controller.Switch(34)=0
  End If
  If vpmKeyUp(keycode) then Exit Sub
End Sub

Sub GetDiff()
  Diff = Int(0 + Rnd *(30-0+1)/10)
End Sub

Sub Leds_Timer()
 Dim ChgLED
 ChgLED = Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
 End Sub

Sub MGOW_exit()
  If B2SOn Then
    Controller.Pause = False
    Controller.Stop
  End If
End Sub

'******************************
'  Setup Desktop
'******************************
Sub Backdrop_Init
  Dim bdl
  If DesktopMode = True then
    LeftSideRail.Visible = 1
    RightSideRail.Visible = 1

    For each bdl in backdrop_lights: bdl.visible = true:Next
    Flasher1.height = 123
    Flasher2.height = 122
    Flasher3.height = 121
    Flasher4.height = 120
    Flasher5.height = 119
    Flasher6.height = 118
    Flasher7.height = 117
    Flasher8.height = 116
    Flasher9.height = 115
    Flasher10.height = 114
    Flasher11.height = 113
    Flasher12.height = 112
  Else
    LeftSideRail.Visible = 0
    RightSideRail.Visible = 0

    For each bdl in backdrop_lights: bdl.visible = false:Next
    Flasher1.height = 143
    Flasher2.height = 142
    Flasher3.height = 141
    Flasher4.height = 140
    Flasher5.height = 139
    Flasher6.height = 138
    Flasher7.height = 137
    Flasher8.height = 136
    Flasher9.height = 135
    Flasher10.height = 134
    Flasher11.height = 133
    Flasher12.height = 132
  End If
End Sub

'''' Apron instructions color
Sub SetInstructions()
  If Apron_Insert_Color = 1 then
'   ApronInsert.image = "apron_insert_black"
    pIC_Left.image = "IC_Left_Black_texture"
    pIC_Right.image = "IC_Right_Black_texture"
  End If

  If Apron_Insert_Color = 2 then
'   ApronInsert.image = "apron_insert_brown"
    pIC_Left.image = "IC_Left_Brown_texture"
    pIC_Right.image = "IC_Right_Brown_texture"
  End If
End Sub

'Apply Cheaters
Sub SetCheats()
  If ROLCheat = 1 Then
    Primitive22.Visible = True
    Rubber5.Collidable = True
    Rubber5.visible = True
  Else
    Primitive22.Visible = False
    Rubber5.Collidable = False
    Rubber5.visible = False
  End If
End Sub

'Set Captive Light Color
Sub SetCaptiveLightColor()
  If CaptiveLightR = 1 then lhalo7b.imageA = "haloBlue":lhalo7b.imageB = "haloBlue" End If
  If CaptiveLightR = 2 then lhalo7b.imageA = "HaloOrange":lhalo7b.imageB = "HaloOrange" End If
  If CaptiveLightR = 3 then lhalo7b.imageA = "haloGreen":lhalo7b.imageB = "haloGreen" End If
  If CaptiveLightR = 4 then lhalo7b.imageA = "haloPurple":lhalo7b.imageB = "haloPurple" End If
  If CaptiveLightR = 5 then lhalo7b.imageA = "haloRed":lhalo7b.imageB = "haloRed" End If
  If CaptiveLightR = 6 then lhalo7b.imageA = "haloWhite":lhalo7b.imageB = "haloWhite" End If
  If CaptiveLightR = 7 then lhalo7b.imageA = "haloYellow":lhalo7b.imageB = "haloYellow" End If

  If CaptiveLightL = 1 then lhalo6b.imageA = "haloBlue":lhalo6b.imageB = "haloBlue" End If
  If CaptiveLightL = 2 then lhalo6b.imageA = "HaloOrange":lhalo6b.imageB = "HaloOrange" End If
  If CaptiveLightL = 3 then lhalo6b.imageA = "haloGreen":lhalo6b.imageB = "haloGreen" End If
  If CaptiveLightL = 4 then lhalo6b.imageA = "haloPurple":lhalo6b.imageB = "haloPurple" End If
  If CaptiveLightL = 5 then lhalo6b.imageA = "haloRed":lhalo6b.imageB = "haloRed" End If
  If CaptiveLightL = 6 then lhalo6b.imageA = "haloWhite":lhalo6b.imageB = "haloWhite" End If
  If CaptiveLightL = 7 then lhalo6b.imageA = "haloYellow":lhalo6b.imageB = "haloYellow" End If
End Sub

''''''''''''''''''''
'Day or Night??
''''''''''''''''''''
Sub CheckDayNight()

End Sub


Sub SetClearPlastics()
  If ClearPlastics = 1 Then
    pPlasticBottomLefta1.visible = False
    pPlasticBottomLeftb1.visible = False
    pPlasticCenter1.visible = False
    pPlasticBottomRight1.visible = False
    pPlasticLowerRight1.visible = False
    pPlasticMidLeft1.visible = False
    pPlasticTop1.visible = False
    pPlasticTopMidRight1.visible = False
    pPlasticTopTriangle1.visible = False
    pPlasticWarBase1.visible = False
  Else
    pPlasticBottomLefta1.visible = True
    pPlasticBottomLeftb1.visible = True
    pPlasticCenter1.visible = True
    pPlasticBottomRight1.visible = True
    pPlasticLowerRight1.visible = True
    pPlasticMidLeft1.visible = True
    pPlasticTop1.visible = True
    pPlasticTopMidRight1.visible = True
    pPlasticTopTriangle1.visible = True
    pPlasticWarBase1.visible = True
End If

End Sub

'Sets GI color
Sub SetGIColor()
  If GI_Color = 2 then
    gi1.color = rgb(255,0,0)
    gi2.color = rgb(255,0,0)
    gi3.color = rgb(255,0,0)
    gi4.color = rgb(255,0,0)
    gi5.color = rgb(255,0,0)
    gi6.color = rgb(255,0,0)
    gi6dtc1.color = rgb(255,0,0)
    gi6dtc2.color = rgb(255,0,0)
    gi6dtc3.color = rgb(255,0,0)
    gi6dtc4.color = rgb(255,0,0)
    gi7.color = rgb(255,255,128)
    gi7dtr1.color = rgb(255,255,128)
    gi7dtr2.color = rgb(255,255,128)
'   gi7dtr3.color = rgb(255,255,128)
    gi7dtr4.color = rgb(255,255,128)
    gi8.color = rgb(255,0,0)
    gi8dtr1.color = rgb(255,0,0)
    gi8dtr2.color = rgb(255,0,0)
    gi8dtr3.color = rgb(255,0,0)
    gi8dtr4.color = rgb(255,0,0)
    gi9.color = rgb(255,0,0)
    gi10.color = rgb(255,0,0)
    gi11.color = rgb(255,0,0)
    gi12.color = rgb(255,0,0)
'   gi13.color = rgb(255,0,0)
    gi14.color = rgb(255,0,0)
    gi15.color = rgb(255,0,0)
    gi16.color = rgb(255,0,0)
    gi17.color = rgb(255,0,0)
    gi18.color = rgb(255,0,0)
    gi19.color = rgb(255,0,0)
    gi20.color = rgb(255,0,0)
    gi21.color = rgb(255,0,0)
    gi21dtr1.color = rgb(255,0,0)
    gi21dtr2.color = rgb(255,0,0)
    gi21dtr3.color = rgb(255,0,0)
    gi21dtr4.color = rgb(255,0,0)
    gi22.color = rgb(255,0,0)
    gi23.color = rgb(255,0,0)
'   gi24.color = rgb(255,0,0)
'   gi25.color = rgb(255,0,0)

    gi1.colorfull = rgb(255,0,0)
    gi2.colorfull = rgb(255,0,0)
    gi3.colorfull = rgb(255,0,0)
    gi4.colorfull = rgb(255,0,0)
    gi5.colorfull = rgb(255,0,0)
    gi6.colorfull = rgb(255,0,0)
    gi6dtc1.colorfull = rgb(255,0,0)
    gi6dtc2.colorfull = rgb(255,0,0)
    gi6dtc3.colorfull = rgb(255,0,0)
    gi6dtc4.colorfull = rgb(255,0,0)
    gi7.colorfull = rgb(255,255,255)
    gi7dtr1.colorfull = rgb(255,255,255)
    gi7dtr2.colorfull = rgb(255,255,255)
'   gi7dtr3.colorfull = rgb(255,255,255)
    gi7dtr4.colorfull = rgb(255,255,255)
    gi8.colorfull = rgb(255,0,0)
    gi8dtr1.colorfull = rgb(255,0,0)
    gi8dtr2.colorfull = rgb(255,0,0)
    gi8dtr3.colorfull = rgb(255,0,0)
    gi8dtr4.colorfull = rgb(255,0,0)
    gi9.colorfull = rgb(255,0,0)
    gi10.colorfull = rgb(255,0,0)
    gi11.colorfull = rgb(255,0,0)
    gi12.colorfull = rgb(255,0,0)
'   gi13.colorfull = rgb(255,0,0)
    gi14.colorfull = rgb(255,0,0)
    gi15.colorfull = rgb(255,0,0)
    gi16.colorfull = rgb(255,0,0)
    gi17.colorfull = rgb(255,0,0)
    gi18.colorfull = rgb(255,0,0)
    gi19.colorfull = rgb(255,0,0)
    gi20.colorfull = rgb(255,0,0)
    gi21.colorfull = rgb(255,0,0)
    gi21dtr1.colorfull = rgb(255,0,0)
    gi21dtr2.colorfull = rgb(255,0,0)
    gi21dtr3.colorfull = rgb(255,0,0)
    gi21dtr4.colorfull = rgb(255,0,0)
    gi22.colorfull = rgb(255,0,0)
    gi23.colorfull = rgb(255,0,0)
'   gi24.colorfull = rgb(255,0,0)

  End If

  If GI_Color = 1 then
    gi1.color = rgb(255,255,128)
    gi2.color = rgb(255,255,128)
    gi3.color = rgb(255,255,128)
    gi4.color = rgb(255,255,128)
    gi5.color = rgb(255,255,128)
    gi6.color = rgb(255,255,128)
    gi6dtc1.color = rgb(255,255,128)
    gi6dtc2.color = rgb(255,255,128)
    gi6dtc3.color = rgb(255,255,128)
    gi6dtc4.color = rgb(255,255,128)
    gi7.color = rgb(255,0,0)
    gi7dtr1.color = rgb(255,0,0)
    gi7dtr2.color = rgb(255,0,0)
'   gi7dtr3.color = rgb(255,0,0)
    gi7dtr4.color = rgb(255,0,0)
    gi8.color = rgb(255,255,128)
    gi8dtr1.color = rgb(255,255,128)
    gi8dtr2.color = rgb(255,255,128)
    gi8dtr3.color = rgb(255,255,128)
    gi8dtr4.color = rgb(255,255,128)
    gi9.color = rgb(255,255,128)
    gi10.color = rgb(255,255,128)
    gi11.color = rgb(255,255,128)
    gi12.color = rgb(255,255,128)
'   gi13.color = rgb(255,255,128)
    gi14.color = rgb(255,255,128)
    gi15.color = rgb(255,255,128)
    gi16.color = rgb(255,255,128)
    gi17.color = rgb(255,255,128)
    gi18.color = rgb(255,255,128)
    gi19.color = rgb(255,255,128)
    gi20.color = rgb(255,255,128)
    gi21.color = rgb(255,255,128)
    gi21dtr1.color = rgb(255,255,128)
    gi21dtr2.color = rgb(255,255,128)
    gi21dtr3.color = rgb(255,255,128)
    gi21dtr4.color = rgb(255,255,128)
    gi22.color = rgb(255,255,128)
    gi23.color = rgb(255,255,128)
'   gi24.color = rgb(255,255,128)
    gi1.colorfull = rgb(255,255,255)
    gi2.colorfull = rgb(255,255,255)
    gi3.colorfull = rgb(255,255,255)
    gi4.colorfull = rgb(255,255,255)
    gi5.colorfull = rgb(255,255,255)
    gi6.colorfull = rgb(255,255,255)
    gi6dtc1.colorfull = rgb(255,255,255)
    gi6dtc2.colorfull = rgb(255,255,255)
    gi6dtc3.colorfull = rgb(255,255,255)
    gi6dtc4.colorfull = rgb(255,255,255)
    gi7.colorfull = rgb(255,0,0)
    gi7dtr1.colorfull = rgb(255,0,0)
    gi7dtr2.colorfull = rgb(255,0,0)
'   gi7dtr3.colorfull = rgb(255,0,0)
    gi7dtr4.colorfull = rgb(255,0,0)
    gi8.colorfull = rgb(255,255,255)
    gi8dtr1.colorfull = rgb(255,255,255)
    gi8dtr2.colorfull = rgb(255,255,255)
    gi8dtr3.colorfull = rgb(255,255,255)
    gi8dtr4.colorfull = rgb(255,255,255)
    gi9.colorfull = rgb(255,255,255)
    gi10.colorfull = rgb(255,255,255)
    gi11.colorfull = rgb(255,255,255)
    gi12.colorfull = rgb(255,255,255)
'   gi13.colorfull = rgb(255,255,255)
    gi14.colorfull = rgb(255,255,255)
    gi15.colorfull = rgb(255,255,255)
    gi16.colorfull = rgb(255,255,255)
    gi17.colorfull = rgb(255,255,255)
    gi18.colorfull = rgb(255,255,255)
    gi19.colorfull = rgb(255,255,255)
    gi20.colorfull = rgb(255,255,255)
    gi21.colorfull = rgb(255,255,255)
    gi21dtr1.colorfull = rgb(255,255,255)
    gi21dtr2.colorfull = rgb(255,255,255)
    gi21dtr3.colorfull = rgb(255,255,255)
    gi21dtr4.colorfull = rgb(255,255,255)
    gi22.colorfull = rgb(255,255,255)
    gi23.colorfull = rgb(255,255,255)
'   gi24.colorfull = rgb(255,255,255)
  End If

End Sub

'''''''''Set Flipper Color

Sub SetFlipperColors()

If Flipper_Color = 1 then
  pLeftFlipper.image = "flipper_red_left"
  pLeftFlipper2.image = "flipper_red_left"
  pRightFlipper.image = "flipper_red_right"
  pRightFlipper2.image = "flipper_red_right"
Else
  pLeftFlipper.image = "flipper_black_left"
  pLeftFlipper2.image = "flipper_black_left"
  pRightFlipper.image = "flipper_black_right"
  pRightFlipper2.image = "flipper_black_right"
End IF

End Sub

'Primitive flippers

Sub Timer_Timer()
  pLeftFlipper.Roty = LeftFlipper.Currentangle
  pLeftFlipper2.Roty = LeftFlipper2.Currentangle
  pRightFlipper.Roty = RightFlipper.Currentangle
  pRightFlipper2.Roty = RightFlipper2.Currentangle
    p_gate1.Rotx = gate1.CurrentAngle + 120

  If Int(gate.CurrentAngle) > 30 Then
    pV_gate.RotZ = 30
  Else
    pV_gate.RotZ = gate.CurrentAngle' + 90
  End If

  If Int(gate2.CurrentAngle) > 30 Then
    pV_gate2.RotZ = 30
  Else
    pV_gate2.RotZ = gate2.CurrentAngle' + 90
  End If

'    pV_gate2.RotZ = gate2.CurrentAngle' + 90
End Sub

'Tube Flasher
Dim TubeFlasherStep

Sub TubeFlasher_Timer
    Select Case TubeFlasherStep
    Case 0:
    Case 1:
    Case 2:
    Case 3:
    Case 4:
    Case 5:
    Case 6:
    Case 7:
    Case 8:
    Case 9:
    Case 10:
        Case 11:SetFlash 101, 1:SetFlash 102, 1:SetFlash 111, 0:SetFlash 112, 0:SetLamp 101, 1:SetLamp 102, 1:SetLamp 111, 0:SetLamp 112, 0
        Case 12:SetFlash 103, 1:SetFlash 104, 1:SetFlash 101, 0:SetFlash 102, 0:SetLamp 103, 1:SetLamp 104, 1:SetLamp 101, 0:SetLamp 102, 0
        Case 13:SetFlash 105, 1:SetFlash 106, 1:SetFlash 103, 0:SetFlash 104, 0:SetLamp 105, 1:SetLamp 106, 1:SetLamp 103, 0:SetLamp 104, 0
        Case 14:SetFlash 107, 1:SetFlash 108, 1:SetFlash 105, 0:SetFlash 106, 0:SetLamp 107, 1:SetLamp 108, 1:SetLamp 105, 0:SetLamp 106, 0
        Case 15:SetFlash 109, 1:SetFlash 110, 1:SetFlash 107, 0:SetFlash 108, 0:SetLamp 109, 1:SetLamp 110, 1:SetLamp 107, 0:SetLamp 108, 0
    Case 16:SetFlash 111, 1:SetFlash 112, 1:SetFlash 109, 0:SetFlash 110, 0:SetLamp 111, 1:SetLamp 112, 1:SetLamp 109, 0:SetLamp 110, 0:TubeFlasherStep = 10
    End Select

    TubeFlasherStep = TubeFlasherStep + 1
End Sub

Sub TubeFlasherOff()
TubeFlasher.enabled = 0
TubeFlasherStep = 0
SetFlash 111, 0
SetFlash 112, 0
SetLamp 111, 0
SetLamp 112, 0
SetFlash 101, 0
SetFlash 102, 0
SetLamp 101, 0
SetLamp 102, 0
SetFlash 103, 0
SetFlash 104, 0
SetLamp 103, 0
SetLamp 104, 0
SetFlash 105, 0
SetFlash 106, 0
SetLamp 105, 0
SetLamp 106, 0
SetFlash 107, 0
SetFlash 108, 0
SetLamp 107, 0
SetLamp 108, 0
SetFlash 109, 0
SetFlash 110, 0
SetLamp 109, 0
SetLamp 110, 0
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''''''
'''SpinnerAngle
'''''''''''''''''''''''''''''''''''''
Dim SpinnerAngle, SpinnerRotations

Sub SW54_Spin
  vpmtimer.pulsesw 54
  PlaySoundAtVol "fx_spinner", sw54, VolSpin
End Sub

Const PI = 3.14

Sub CheckSpinnerRod_timer()
  pSpinnerRod.TransX = sin( (SW54.CurrentAngle+180) * (2*PI/360)) * 5
  pSpinnerRod.TransY = sin( (SW54.CurrentAngle- 90) * (2*PI/360)) * 5
End Sub

' drop targets
Dim iauStep, iauCenter, iauRight


'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'   Center
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

Sub DC1_Hit:vpmTimer.PulseSw 22:DC1.IsDropped = True:gi6dtc1.state = 1:Me.TimerEnabled = 1:PlaySoundAtVol SoundFX("droptarget",DOFDropTargets),ActiveBall,1:End Sub
Sub DC1_Timer:Me.TimerEnabled = 0 'This Drops The Target
End Sub

Sub DC2_Hit:vpmTimer.PulseSw 32:DC2.IsDropped = True:gi6dtc2.state = 1:Me.TimerEnabled = 1:PlaySoundAtVol SoundFX("droptarget",DOFDropTargets),ActiveBall, 1:End Sub
Sub DC2_Timer:Me.TimerEnabled = 0 'This Drops The Target
End Sub

Sub DC3_Hit:vpmTimer.PulseSw 42:DC3.IsDropped = True:gi6dtc3.state = 1:Me.TimerEnabled = 1:PlaySoundAtVol SoundFX("droptarget",DOFDropTargets),ActiveBall, 1:End Sub
Sub DC3_Timer:Me.TimerEnabled = 0 'This Drops The Target
End Sub

Sub DC4_Hit:vpmTimer.PulseSw 52:DC4.IsDropped = True:gi6dtc4.state = 1:Me.TimerEnabled = 1:PlaySoundAtVol SoundFX("droptarget",DOFDropTargets),ActiveBall, 1:End Sub
Sub DC4_Timer:Me.TimerEnabled = 0 'This Drops The Target
End Sub

Sub ResetCDrops(Enabled)
  If iauCenter = 1 Then
    ImAlreadyUp.Enabled = True
  Else
  iauCenter = 1
  PlaySoundAtVol SoundFX("mgow_Center_DropsUp",DOFContactors), DC2, 1
  DC1.IsDropped=False
  gi6dtc1.state = 0
  DC2.IsDropped=False
  gi6dtc2.state = 0
  DC3.IsDropped=False
  gi6dtc3.state = 0
  DC4.IsDropped=False
  gi6dtc4.state = 0
  End If
End Sub

'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'   Top Right
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

Sub DR1_Hit:vpmTimer.PulseSw 21:DR1.IsDropped=True:gi8dtr1.state = 1:gi7dtr1.state = 1:gi21dtr1.state = 1:Me.TimerEnabled = 1:PlaySoundAtVol SoundFX("droptarget",DOFDropTargets), ActiveBall, 1:End Sub
Sub DR1_Timer:Me.TimerEnabled = 0: 'This Drops The Target
End Sub

Sub DR2_Hit:vpmTimer.PulseSw 31:DR2.IsDropped=True:gi8dtr2.state = 1:gi7dtr2.state = 1:gi21dtr2.state = 1:Me.TimerEnabled = 1:PlaySoundAtVol SoundFX("droptarget",DOFDropTargets), ActiveBall, 1:End Sub
Sub DR2_Timer:Me.TimerEnabled = 0: 'This Drops The Target
End Sub

Sub DR3_Hit:vpmTimer.PulseSw 41:DR3.IsDropped=True:gi8dtr3.state = 1::gi21dtr3.state = 1:Me.TimerEnabled = 1:PlaySoundAtVol SoundFX("droptarget",DOFDropTargets), Activeball, 1:End Sub 'gi7dtr3.state = 1
Sub DR3_Timer:Me.TimerEnabled = 0: 'This Drops The Target
End Sub

Sub DR4_Hit:vpmTimer.PulseSw 51:DR4.IsDropped=True:gi8dtr4.state = 1:gi7dtr4.state = 1:gi21dtr4.state = 1:Me.TimerEnabled = 1:PlaySoundAtVol SoundFX("droptarget",DOFDropTargets), Activeball, 1:End Sub
Sub DR4_Timer:Me.TimerEnabled = 0: 'This Drops The Target
End Sub

Sub ResetRDrops(Enabled)
  If iauRight = 1 Then
    ImAlreadyUp.Enabled = True
  Else
  iauRight = 1
  PlaySoundAtVol SoundFX("mgow_Right_DropsUp",DOFContactors), gi7dtr1, VolTarg
  DR1.IsDropped=False:
  gi8dtr1.state = 0
  gi7dtr1.state = 0
  gi21dtr1.state = 0
  DR2.IsDropped=False:
  gi8dtr2.state = 0
  gi7dtr2.state = 0
  gi21dtr2.state = 0
  DR3.IsDropped=False:
  gi8dtr3.state = 0
' gi7dtr3.state = 0
  gi21dtr3.state = 0
  DR4.IsDropped=False:
  gi8dtr4.state = 0
  gi7dtr4.state = 0
  gi21dtr4.state = 0
  End If

End Sub

'targets always reset twice..  checks to see if they are already up then makes sure they don't go up again

Sub ImAlreadyUp_Timer()
           Select Case iauStep
               Case 1:
               Case 2:
               Case 3:
               Case 4:
               Case 5:
               Case 6:
               Case 7:
               Case 8:ImAlreadyUp.Enabled = False:iauRight = 0:iauCenter = 0:iauStep = 0
           End Select
    iauStep = iauStep + 1
       End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''Bumpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim bump1,bump2,bump3,bump4,bump5,bump6
Sub Bumper1_Hit:vpmTimer.PulseSw(23):bump1 = 1:Me.TimerEnabled = 1:PlaySoundAtVol SoundFX("fx_bumper1",DOFContactors), Bumper1, VolBump:End Sub
Sub Bumper2_Hit:vpmTimer.PulseSw(33):bump2 = 1:Me.TimerEnabled = 1:PlaySoundAtVol SoundFX("fx_bumper2",DOFContactors), Bumper2, VolBump:End Sub
Sub Bumper3_Hit:vpmTimer.PulseSw(43):bump3 = 1:Me.TimerEnabled = 1:PlaySoundAtVol SoundFX("fx_bumper3",DOFContactors), Bumper3, VolBump:End Sub
Sub Bumper4_Hit:vpmTimer.PulseSw(53):bump4 = 1:Me.TimerEnabled = 1:PlaySoundAtVol SoundFX("fx_bumper4",DOFContactors), Bumper4, VolBump:End Sub

       Sub Bumper1_Timer()
           Select Case bump1
               Case 1:'BR1.z = 0:
               Case 2:'BR1.z = -15:
               Case 3:'BR1.z = -30:
               Case 4:'BR1.z = -40:
               Case 5:'BR1.z = -30:
               Case 6:'BR1.z = -15:
               Case 7:'BR1.z = 0:
               Case 8::Me.TimerEnabled = 0:'BR1.z = 10
           End Select
    Bump1 = bump1 + 1
       End Sub

       Sub Bumper2_Timer()
           Select Case bump2
               Case 1:'BR2.z = 0:
               Case 2:'BR2.z = -15:
               Case 3:'BR2.z = -30:
               Case 4:'BR2.z = -40:
               Case 5:'BR2.z = -30:
               Case 6:'BR2.z = -15:
               Case 7:'BR2.z = 0:
               Case 8:Me.TimerEnabled = 0:'BR2.z = 10
           End Select
    Bump2 = bump2 + 1
       End Sub

       Sub Bumper3_Timer()
           Select Case bump3
               Case 1:'BR3.z = 0:
               Case 2:'BR3.z = -15:
               Case 3:'BR3.z = -30:
               Case 4:'BR3.z = -40:
               Case 5:'BR3.z = -30:
               Case 6:'BR3.z = -15:
               Case 7:'BR3.z = 0:
               Case 8:Me.TimerEnabled = 0:'BR3.z = 10
           End Select
    Bump3 = bump3 + 1
       End Sub

       Sub Bumper4_Timer()
           Select Case bump4
               Case 1:'BR4.z = 0:
               Case 2:'BR4.z = -15:
               Case 3:'BR4.z = -30:
               Case 4:'BR4.z = -40:
               Case 5:'BR4.z = -30:
               Case 6:'BR4.z = -15:
               Case 7:'BR4.z = 00:
               Case 8:Me.TimerEnabled = 0:'BR4.z = 10
           End Select
    Bump4 = bump4 + 1
       End Sub

''''''''''''''''''''''''''''''''''''
' Warebace kickers
''''''''''''''''''''''''''''''''''''
dim SW61Step, SW62Step

Sub SW61_Timer()
  Select Case SW61Step
    Case 0: TWKicker1.transY = 2
    Case 1: TWKicker1.transY = 5
    Case 2: TWKicker1.transY = 5
    Case 3: TWKicker1.transY = 2
    Case 4: TWKicker1.transY = 0:Me.TimerEnabled = 0:SW61Step = 0
  End Select
  SW61Step = SW61Step + 1
End Sub

Sub SW62_Timer()
  Select Case SW62Step
    Case 0: TWKicker2.transY = 2
    Case 1: TWKicker2.transY = 5
    Case 2: TWKicker2.transY = 5
    Case 3: TWKicker2.transY = 2
    Case 4: TWKicker2.transY = 0:Me.TimerEnabled = 0:SW62Step = 0
  End Select
  SW62Step = SW62Step + 1
End Sub

'delayed sound timer for launch (played many sounds without)
Dim SW63Step
Sub SW63_Timer()
    Select Case SW63Step
    Case 0:
    Case 1:pKickRod.ObjRotX = 10:PlaySoundAtVol SoundFX("mgow kicker out",DOFContactors),pKickRod,VolKick:SW63.Kick 0,15 + diff
    Case 2:pKickRod.ObjRotX = 5:Controller.Switch(63)=0
    Case 3:pKickRod.ObjRotX = 0:
    Case 4:pKickRod.ObjRotX = -5:
    Case 5:pKickRod.ObjRotX = 0:
    Case 6:
    Case 7:
    Case 8:
    Case 9:
    Case 10:me.timerenabled = 0 :SW63Step = 0:sw63active = 0
    End Select

    SW63Step = SW63Step + 1
End Sub

''''''''''''''''''''''''''''''''''
' Yellow Target
''''''''''''''''''''''''''''''''''
Dim Tsw71Step

Sub Tsw71_Hit:vpmTimer.PulseSw(71):PTarget1.TransX = -5:Tsw71Step = 1:PlaySoundAtVol SoundFX("fx_target",DOFTargets),PTarget1,VolTarg:Me.TimerEnabled = 1:End Sub
Sub Tsw71_timer()
  Select Case Tsw71Step
    Case 1:PTarget1.TransX = 3
        Case 2:PTarget1.TransX = -2
        Case 3:PTarget1.TransX = 1
        Case 4:PTarget1.TransX = 0:Me.TimerEnabled = 0
     End Select
  Tsw71Step = Tsw71Step + 1
End Sub

''''''''''''''''''''''''''''''''''''''
'Roll over switches
''''''''''''''''''''''''''''''''''''''
Sub SW60_Hit:Controller.Switch(60)=1:TWsw60.transY = -15:PlaySoundAtVol "rollover", Activeball, 1:End Sub
Sub SW60_unHit:Controller.Switch(60)=0:TWsw60.transY = 0:End Sub

Sub SW30_Hit:Controller.Switch(30)=1:TWsw30.transY = -15:PlaySoundAtVol "rollover", Activeball, 1:DOF 101, DOFOn:End Sub
Sub SW30_unHit:Controller.Switch(30)=0:TWsw30.transY = 0:DOF 101, DOFOff:End Sub

Sub SW72_Hit:Controller.Switch(72)=1:TWsw72.transY = -15:PlaySoundAtVol "rollover", Activeball, 1:End Sub
Sub SW72_unHit:Controller.Switch(72)=0:TWsw72.transY = 0:End Sub

Sub SW55_Hit:Controller.Switch(55)=1:TWsw55.transY = -15:PlaySoundAtVol "rollover", Activeball, 1:End Sub
Sub SW55_unHit:Controller.Switch(55)=0:TWsw55.transY = 0:End Sub

Sub SW70_Hit:Controller.Switch(70)=1:TWsw70.transY = -15:PlaySoundAtVol "rollover", Activeball, 1:End Sub
Sub SW70_unHit:Controller.Switch(70)=0:TWsw70.transY = 0:End Sub

Sub SW20_Hit:Controller.Switch(20)=1:TWsw20.transY = -15:PlaySoundAtVol "rollover", Activeball, 1:End Sub
Sub SW20_unHit:Controller.Switch(20)=0:TWsw20.transY = 0:End Sub

Sub SW30a_Hit:Controller.Switch(30)=1:TWsw30a.transY = -15:PlaySoundAtVol "rollover", Activeball, 1:DOF 102, DOFOn:End Sub
Sub SW30a_unHit:Controller.Switch(30)=0:TWsw30a.transY = 0:DOF 102, DOFOff:End Sub

Sub SW40_Hit:Controller.Switch(40)=1:TWsw40.transY = -15:PlaySoundAtVol "rollover", Activeball, 1:End Sub
Sub SW40_unHit:Controller.Switch(40)=0:TWsw40.transY = 0:End Sub

Sub SW50_Hit:Controller.Switch(50)=1:TWsw50.transY = -15:PlaySoundAtVol "rollover", Activeball, 1:End Sub
Sub SW50_unHit:Controller.Switch(50)=0:TWsw50.transY = 0:End Sub


''''''''''''''''''''''''''''''''''''''''
'Slings
''''''''''''''''''''''''''''''''''''''''

Dim Sling1Step, Sling2Step
Sub Sling1_slingshot:vpmTimer.PulseSw(73):PlaySoundAtVol SoundFXDOF("left_slingshot",103,DOFPulse,DOFContactors), Activeball, 1:Sling1a.visible = false:pSling1.TransZ = -10:Sling1b.visible = true:Sling1Step = 0:Me.TimerEnabled = 1:DOF 103, 2:End Sub

Sub Sling1_Timer
    Select Case Sling1Step
        Case 0:Sling1b.visible = false:pSling1.TransZ = -20:Sling1c.visible = true:
        Case 1:Sling1c.visible = false:pSling1.TransZ = -30:Sling1d.visible = true:
        Case 2:Sling1d.visible = false:pSling1.TransZ = -20:Sling1c.visible = true:
        Case 3:Sling1c.visible = false:pSling1.TransZ = -10:Sling1b.visible = true:
        Case 4:Sling1b.visible = false:pSling1.TransZ = 0:Sling1a.visible = true::Me.TimerEnabled = 0 '
    End Select

    Sling1Step = Sling1Step + 1
End Sub

Sub Sling2_slingshot:vpmTimer.PulseSw(73):PlaySoundAtVol SoundFXDOF("Right_slingshot",104,DOFPulse,DOFContactors), Activeball, 1:Sling2a.visible = false:pSling2.TransZ = -10:Sling2b.visible = true:Sling2Step = 0:Me.TimerEnabled = 1:End Sub

Sub Sling2_Timer
    Select Case Sling2Step
        Case 0:Sling2b.visible = false:pSling2.TransZ = -20:Sling2c.visible = true:
        Case 1:Sling2c.visible = false:pSling2.TransZ = -30:Sling2d.visible = true:
        Case 2:Sling2d.visible = false:pSling2.TransZ = -20:Sling2c.visible = true:
        Case 3:Sling2c.visible = false:pSling2.TransZ = -10:Sling2b.visible = true:
        Case 4:Sling2b.visible = false:pSling2.TransZ = 0:Sling2a.visible = true::Me.TimerEnabled = 0 '
    End Select

    Sling2Step = Sling2Step + 1
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''  Ball Through system''''''''''''''''''''''''''
'''''''''''''''''''''by cyberpez''''''''''''''''''''''''''''''''
''''''''''''''''based off of EalaDubhSidhe's''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim BallCount

Sub CheckMaxBalls()
  BallCount = MaxBalls
End Sub

Sub CreatBalls()
  If BallCount > 0 then
    drain.CreateSizedBallWithMass BallRadius, BallMass
'   Drain.CreateBall
    Drain.kick 70,10
    BallCount = BallCount - 1
  End If
End Sub

dim bstatus

Sub CheckBallStatus_timer()

  Select Case Bstatus
  Case 1: If Kicker1active = 0 and Kicker2active = 1 then Kicker2.Kick 70,10 End If
  Case 2: If Kicker2active = 0 and Kicker3active = 1 then Kicker3.Kick 70,10 End If
  Case 3: If Kicker3active = 0 and Kicker4active = 1 then Kicker4.Kick 70,10 End If
  Case 4: If Kicker4active = 0 and Kicker5active = 1 then Kicker5.Kick 70,10 End If
  Case 5: If Kicker5active = 0 and Kicker6active = 1 then Kicker6.Kick 70,10 End If
  Case 6: bstatus = 0
  End Select
  bstatus = bstatus + 1
End Sub

Dim Kicker1active, Kicker2active, Kicker3active, Kicker4active, Kicker5active, Kicker6active

Sub Kicker1_hit()
  Kicker1active = 1
' Controller.Switch(13)=1
End Sub

Sub Kicker1_unhit()

End Sub

Sub Kicker2_hit()
  Kicker2active = 1
End Sub

Sub Kicker2_unhit()
  Kicker2active = 0
End Sub

Sub Kicker3_hit()
  Kicker3active = 1
  Controller.Switch(44)=1
End Sub

Sub Kicker3_unhit()
  Kicker3active = 0
  Controller.Switch(44)=0
End Sub

Sub Kicker4_hit()
  Kicker4active = 1
End Sub

Sub Kicker4_unhit()
  Kicker4active = 0
End Sub

Sub Kicker5_hit()
  Kicker5active = 1
    controller.switch(64) = false
End Sub

Sub Kicker5_unhit()
  Kicker5active = 0
End Sub

Sub Kicker6_hit()
  Kicker6active = 1
End Sub

Sub Kicker6_unhit()
  Kicker6active = 0
    If BallCount > 0 then
      CreatBalls
    End If
End Sub

dim DontKickAnyMoreBalls

Sub KickBallToLane()
  If DontKickAnyMoreBalls = 0 then
    Kicker1.Kick 70,5
    bstatus = 2
    Kicker1active = 0
    iBall = iBall - 1
    fgBall = false
    DontKickAnyMoreBalls = 1
    DKTMstep = 1
    DontKickToMany.enabled = true
    BallsInPlay = BallsInPlay + 1
  End If
End Sub

Dim DKTMstep

Sub DontKickToMany_timer ()
  Select Case DKTMstep
  Case 1:
  Case 2:
  Case 3: DontKickAnyMoreBalls = 0:DontKickToMany.Enabled = False: DontKickAnyMoreBalls = 0
  End Select
  DKTMstep = DKTMstep + 1
End Sub

sub kisort()
  if fgBall then
    Drain.Kick 70,10
    iBall = iBall + 1
    fgBall = false
  end if
end sub


Sub Drain_hit()
  PlaySoundAtVol "drain", drain ,1
  controller.switch(64) = true
  fgBall = true
  iBall = iBall + 1
  BallsInPlay = BallsInPlay - 1
End Sub

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'     LL    EEEEEE  DDDD    ,,   SSSSS
'         LL    EE    DD  DD    ,,  SS
'     LL    EE    DD   DD    ,   SS
'     LL    EEEE  DD   DD        SS
'     LL    EE    DD  DD          SS
'     LLLLLL  EEEEEE  DDDD      SSSSS
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
'   6 Didget Array
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
Dim LED(35)
LED(0)=Array(d121,d122,d123,d124,d125,d126,d127,LXM,d128)
LED(1)=Array(d131,d132,d133,d134,d135,d136,d137,LXM,d138)
LED(2)=Array(d141,d142,d143,d144,d145,d146,d147,LXM,d148)
LED(3)=Array(d151,d152,d153,d154,d155,d156,d157,LXM,d158)
LED(4)=Array(d161,d162,d163,d164,d165,d166,d167,LXM,d168)
LED(5)=Array(d171,d172,d173,d174,d175,d176,d177,LXM,d178)

LED(6)=Array(d221,d222,d223,d224,d225,d226,d227,LXM,d228)
LED(7)=Array(d231,d232,d233,d234,d235,d236,d237,LXM,d238)
LED(8)=Array(d241,d242,d243,d244,d245,d246,d247,LXM,d248)
LED(9)=Array(d251,d252,d253,d254,d255,d256,d257,LXM,d258)
LED(10)=Array(d261,d262,d263,d264,d265,d266,d267,LXM,d268)
LED(11)=Array(d271,d272,d273,d274,d275,d276,d277,LXM,d278)

LED(12)=Array(d321,d322,d323,d324,d325,d326,d327,LXM,d328)
LED(13)=Array(d331,d332,d333,d334,d335,d336,d337,LXM,d338)
LED(14)=Array(d341,d342,d343,d344,d345,d346,d347,LXM,d348)
LED(15)=Array(d351,d352,d353,d354,d355,d356,d357,LXM,d358)
LED(16)=Array(d361,d362,d363,d364,d365,d366,d367,LXM,d368)
LED(17)=Array(d371,d372,d373,d374,d375,d376,d377,LXM,d378)

LED(18)=Array(d421,d422,d423,d424,d425,d426,d427,LXM,d428)
LED(19)=Array(d431,d432,d433,d434,d435,d436,d437,LXM,d438)
LED(20)=Array(d441,d442,d443,d444,d445,d446,d447,LXM,d448)
LED(21)=Array(d451,d452,d453,d454,d455,d456,d457,LXM,d458)
LED(22)=Array(d461,d462,d463,d464,d465,d466,d467,LXM,d468)
LED(23)=Array(d471,d472,d473,d474,d475,d476,d477,LXM,d478)

LED(24)=Array(d511,d512,d513,d514,d515,d516,d517,LXM,d518)
LED(25)=Array(d521,d522,d523,d524,d525,d526,d527,LXM,d528)
LED(26)=Array(d611,d612,d613,d614,d615,d616,d617,LXM,d618)
LED(27)=Array(d621,d622,d623,d624,d625,d626,d627,LXM,d628)

Sub DisplayTimer_Timer
Dim ChgLED, ii, num, chg, stat, obj
ChgLED = Controller.ChangedLEDs (&Hffffffff, &Hffffffff)
If Not IsEmpty (ChgLED) Then
For ii = 0 To UBound (chgLED)
num = chgLED (ii, 0) : chg = chgLED (ii, 1) : stat = chgLED (ii, 2)
For Each obj In LED (num)
If chg And 1 Then obj.State = stat And 1
chg = chg \ 2 : stat = stat \ 2
Next
Next
End If
End Sub

'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
'   7 Didget Array
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
Dim LED7(35)
LED7(0)=Array(d111,d112,d113,d114,d115,d116,d117,LXM,d118)
LED7(1)=Array(d121,d122,d123,d124,d125,d126,d127,LXM,d128)
LED7(2)=Array(d131,d132,d133,d134,d135,d136,d137,LXM,d138)
LED7(3)=Array(d141,d142,d143,d144,d145,d146,d147,LXM,d148)
LED7(4)=Array(d151,d152,d153,d154,d155,d156,d157,LXM,d158)
LED7(5)=Array(d161,d162,d163,d164,d165,d166,d167,LXM,d168)
LED7(6)=Array(d171,d172,d173,d174,d175,d176,d177,LXM,d178)

LED7(7)=Array(d211,d212,d213,d214,d215,d216,d217,LXM,d218)
LED7(8)=Array(d221,d222,d223,d224,d225,d226,d227,LXM,d228)
LED7(9)=Array(d231,d232,d233,d234,d235,d236,d237,LXM,d238)
LED7(10)=Array(d241,d242,d243,d244,d245,d246,d247,LXM,d248)
LED7(11)=Array(d251,d252,d253,d254,d255,d256,d257,LXM,d258)
LED7(12)=Array(d261,d262,d263,d264,d265,d266,d267,LXM,d268)
LED7(13)=Array(d271,d272,d273,d274,d275,d276,d277,LXM,d278)

LED7(14)=Array(d311,d312,d313,d314,d315,d316,d317,LXM,d318)
LED7(15)=Array(d321,d322,d323,d324,d325,d326,d327,LXM,d328)
LED7(16)=Array(d331,d332,d333,d334,d335,d336,d337,LXM,d338)
LED7(17)=Array(d341,d342,d343,d344,d345,d346,d347,LXM,d348)
LED7(18)=Array(d351,d352,d353,d354,d355,d356,d357,LXM,d358)
LED7(19)=Array(d361,d362,d363,d364,d365,d366,d367,LXM,d368)
LED7(20)=Array(d371,d372,d373,d374,d375,d376,d377,LXM,d378)

LED7(21)=Array(d411,d412,d413,d414,d415,d416,d417,LXM,d418)
LED7(22)=Array(d421,d422,d423,d424,d425,d426,d427,LXM,d428)
LED7(23)=Array(d431,d432,d433,d434,d435,d436,d437,LXM,d438)
LED7(24)=Array(d441,d442,d443,d444,d445,d446,d447,LXM,d448)
LED7(25)=Array(d451,d452,d453,d454,d455,d456,d457,LXM,d458)
LED7(26)=Array(d461,d462,d463,d464,d465,d466,d467,LXM,d468)
LED7(27)=Array(d471,d472,d473,d474,d475,d476,d477,LXM,d478)

LED7(28)=Array(d511,d512,d513,d514,d515,d516,d517,LXM,d518)   'was 24 -- 26
LED7(29)=Array(d521,d522,d523,d524,d525,d526,d527,LXM,d528)   'was 25 -- 27
LED7(30)=Array(d611,d612,d613,d614,d615,d616,d617,LXM,d618)   'was 26 -- 24
LED7(31)=Array(d621,d622,d623,d624,d625,d626,d627,LXM,d628)   'was 27 -- 25

Sub DisplayTimer7_Timer
Dim ChgLED, ii, num, chg, stat, obj
ChgLED = Controller.ChangedLEDs (&Hffffffff, &Hffffffff)
If Not IsEmpty (ChgLED) Then
For ii = 0 To UBound (chgLED)
num = chgLED (ii, 0) : chg = chgLED (ii, 1) : stat = chgLED (ii, 2)
For Each obj In LED7 (num)
If chg And 1 Then obj.State = stat And 1
chg = chg \ 2 : stat = stat \ 2
Next
Next
End If
End Sub

'***************************************************
'  JP's Fading Lamps & Flashers version 9 for VP921
'   Based on PD's Fading Lights
' SetLamp 0 is Off
' SetLamp 1 is On
' FadingLevel(x) = fading state
' LampState(x) = light state
' Includes the flash element (needs own timer)
' Flashers can be used as lights too
'***************************************************

Dim LampState(200), FadingLevel(200)
Dim FlashState(200), FlashLevel(200)
Dim FlashSpeedUp, FlashSpeedDown
Dim FadeArray1: FadeArray1 = Array("mgof_on", "mgof_66", "mgof_33", "mgof_off")
Dim BulbArray: BulbArray = Array("bulbred_66", "bulbred_66", "bulbred_33", "bulbred_off")
Dim BulbArray2: BulbArray2 = Array("bulb_on", "bulb_66", "bulb_33", "bulb_off")
Dim BArray: BArray = Array("mgow_bumpercap_on", "mgow_bumpercap_66", "mgow_bumpercap_33", "mgow_bumpercap_off")
Dim BArrayNight: BArrayNight = Array("mgow_bumpercap_on", "mgow_bumpercap_66", "mgow_bumpercap_33", "mgow_bumpercap_off_night")
Dim TextureArray1: TextureArray1 = Array("Plastic with an image trans", "Plastic with an image")

Const LightHaloBrightness   = 200

FlashInit()
FlasherTimer.Interval = 10 'flash fading speed
FlasherTimer.Enabled = 1

' Lamp & Flasher Timers

Sub StartLampTimer
  AllLampsOff()
  LampTimer.Interval = 30 'lamp fading speed
  LampTimer.Enabled = 1
End Sub

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
  UpdateMultipleLamps
End Sub

Sub UpdateLamps

'Bumper cap lights
  FadeMaterialP 45, PBumpCap1, TextureArray1
  NFadeLm 45, Blight1a
  NFadeLm 45, Blight1b
  FadeMaterialP 47, PBumpCap2, TextureArray1
  NFadeLm 47, Blight2a
  NFadeLm 47, Blight2b
  FadeMaterialP 44, PBumpCap3, TextureArray1
  NFadeLm 44, blight3a
  NFadeLm 44, Blight3b
  FadeMaterialP 46, PBumpCap4, TextureArray1
  NFadeLm 46, Blight4a
  NFadeLm 46, Blight4b

  FadePri4 101, Bulb1, BulbArray
  FadePri4 102, Bulb2, BulbArray
  FadePri4 103, Bulb3, BulbArray
  FadePri4 104, Bulb4, BulbArray
  FadePri4 105, Bulb5, BulbArray
  FadePri4 106, Bulb6, BulbArray
  FadePri4 107, Bulb7, BulbArray
  FadePri4 108, Bulb8, BulbArray
  FadePri4 109, Bulb9, BulbArray
  FadePri4 110, Bulb10, BulbArray
  FadePri4 111, Bulb11, BulbArray
  FadePri4 112, Bulb12, BulbArray

  NFadeL 3, Light3a
  NFadeL 4, Light4
  NFadeL 5, Light5
  NFadeL 6, Light6
  NFadeL 7, Light7

  NFadeLm 14, Light14a
  NFadeLm 14, Light14b
  NFadeL 14, Light14c
  NFadeLm 15, Light15a
  NFadeLm 15, Light15b
  NFadeL 15, Light15c
  NFadeLm 16, Light16a
  NFadeLm 16, Light16b
  NFadeL 16, Light16c
  NFadeLm 17, Light17a
  NFadeLm 17, Light17b
  NFadeL 17, Light17c
  NFadeLm 18, Light18a
  NFadeL 18, Light18b
  NFadeLm 19, Light19a
  NFadeL 19, Light19b

  NFadeL 20, Light20
  NFadeL 21, Light21
  NFadeLm 22, Light22a
  NFadeLm 23, Light23a
  NFadeL 23, Light23b
  NFadeLm 24, Light24a
  NFadeL 24, Light24b
  NFadeLm 25, Light25a
  NFadeL 25, Light25b
  NFadeLm 26, Light26a
  NFadeL 26, Light26b
  NFadeL 27, Light27
  NFadeL 28, Light28
  NFadeL 29, Light29
  NFadeL 30, Light30

  NFadeL 31, Light31
  NFadeL 32, Light32
  NFadeL 33, Light33
  NFadeL 34, Light34
  NFadeL 35, Light35
  NFadeL 36, Light36
  NFadeL 37, Light37
  NFadeL 38, Light38
  NFadeL 39, Light39
  NFadeL 40, Light40

  NFadeL 41, Light41
  NFadeL 42, Light42
  NFadeL 43, Light43

  NFadeL 48, Light48
  NFadeLm 49, Light49a
  NFadeL 49, Light49b
  NFadeLm 50, Light50a
  NFadeL 50, Light50b

  NFadeL 51, Light51

 End Sub

'Sindbad: You can use this instead of FadeLN
' call it this way: FadeLight lampnumber, light, Array
Sub FadeLight(nr, Light, group)
    Select Case FadingLevel(nr)
        Case 2:Light.image = group(3):FadingLevel(nr) = 0 'Off
        Case 3:Light.image = group(2):FadingLevel(nr) = 2 'fading...
        Case 4:Light.image = group(1):FadingLevel(nr) = 3 'fading...
        Case 5:Light.image = group(0):FadingLevel(nr) = 1 'ON
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


LightHalos_Init
Sub LightHalos_Init
  If LightHalo_ON = 0 Then 'hide halos
    For each xx in aLightHalos:xx.IsVisible = False:Next
    For each xx in aLightHalos:xx.Alpha = 0:Next
  End If



End Sub

Sub FlasherTimer_Timer()

  Flash 101, Flasher1
  Flash 102, Flasher2
  Flash 103, Flasher3
  Flash 104, Flasher4
  Flash 105, Flasher5
  Flash 106, Flasher6
  Flash 107, Flasher7
  Flash 108, Flasher8
  Flash 109, Flasher9
  Flash 110, Flasher10
  Flash 111, Flasher11
  Flash 112, Flasher12

    If CaptiveLightL = 0 then
    Else
    FlashVal 6, lhalo6b, LightHaloBrightness
    End If

    If CaptiveLightR = 0 then
    Else
    FlashVal 7, lhalo7b, LightHaloBrightness
    End If

'***GI
  If LampState(0) = 1 Then SetGIOn
  If LampState(0) = 0 Then SetGIOff


  If LampState(0) = 1 and LampState(22) = 1 then
    NFadeLm 22 , gi7a
    NFadeL 22 , gi7

  End If
  If LampState(0) = 1 and LampState(22) = 0 then
  End If

  NFadeLm 0, gi1
  NFadeLm 0, gi1a
  NFadeLm 0, gi2
  NFadeLm 0, gi2a
  NFadeLm 0, gi3
  NFadeLm 0, gi3a
  NFadeLm 0, gi4
  NFadeLm 0, gi4a
  NFadeLm 0, gi5
  NFadeLm 0, gi5a
  NFadeLm 0, gi6
  NFadeLm 0, gi6a
  NFadeLm 0, gi6b

  NFadeLm 0, gi8
  NFadeLm 0, gi8a
  NFadeLm 0, gi9
  NFadeLm 0, gi9a
  NFadeLm 0, gi10
  NFadeLm 0, gi10a
  NFadeLm 0, gi11
  NFadeLm 0, gi11a
  NFadeLm 0, gi12
  NFadeLm 0, gi12a
' NFadeLm 0, gi13
' NFadeLm 0, gi13a
  NFadeLm 0, gi14
  NFadeLm 0, gi15
  NFadeLm 0, gi16
  NFadeLm 0, gi17
  NFadeLm 0, gi18
  NFadeLm 0, gi19
  NFadeLm 0, gi19a
  NFadeLm 0, gi20
  NFadeLm 0, gi20a
  NFadeLm 0, gi21
  NFadeLm 0, gi21a
  NFadeLm 0, gi22
  NFadeLm 0, gi23
  NFadeL 0, gi23a
' NFadeLm 0, gi24a
' NFadeL 0, gi24

End Sub

'trxture swap
dim itemw, itemp

Sub FadeMaterialW(nr, itemw, group)
    Select Case FadingLevel(nr)
        Case 4:itemw.TopMaterial = group(1):itemw.SideMaterial = group(1)
        Case 5:itemw.TopMaterial = group(0):itemw.SideMaterial = group(0)
    End Select
End Sub

Sub FadeMaterialP(nr, itemp, group)
    Select Case FadingLevel(nr)
        Case 4:itemp.Material = group(1)
        Case 5:itemp.Material = group(0)
    End Select
End Sub

' div lamp subs
Sub AllLampsOff()
    Dim x
    For x = 0 to 200
        LampState(x) = 0
        FadingLevel(x) = 4
    Next

UpdateLamps:UpdateLamps:Updatelamps
End Sub

Sub SetLamp(nr, value)
    If value <> LampState(nr) Then
        LampState(nr) = abs(value)
        FadingLevel(nr) = abs(value) + 4
    End If
End Sub

''''''''''''''''''''''''''''''''''''
'''GGGGG IIIII  OOOOO N   N
'''G       I    O   O NN  N
'''G  GG   I    O   O N N N
'''G   G   I    O   O N  NN
'''GGGGG IIIII  OOOOO N   N
''''''''''''''''''''''''''''''''''''
Sub SetGIOn
  Primitive1.image = "[plastic-red_on]"
  Primitive2.image = "[plastic-red_on]"
  Primitive3.image = "[plastic-red_on]"
  Primitive4.image = "[plastic-red_on]"
  Primitive5.image = "[plastic-red_on]"
End Sub

''''''''''''''''''''''''''''''''''''
'''GGGGG IIIII  OOOOO FFFFF FFFFF
'''G       I    O   O F     F
'''G  GG   I    O   O FFF   FFF
'''G   G   I    O   O F     F
'''GGGGG IIIII  OOOOO F     F
''''''''''''''''''''''''''''''''''''
Sub SetGIOff

  Primitive1.image = "[plastic-red]"
  Primitive2.image = "[plastic-red]"
  Primitive3.image = "[plastic-red]"
  Primitive4.image = "[plastic-red]"
  Primitive5.image = "[plastic-red]"
End Sub

' div flasher subs
Sub FlashInit
    Dim i
    For i = 0 to 200
        FlashState(i) = 0
        FlashLevel(i) = 0
    Next

    FlashSpeedUp = 30   ' fast speed when turning on the flasher
    FlashSpeedDown = 10 ' slow speed when turning off the flasher, gives a smooth fading
    ' you could also change the default images for each flasher or leave it as in the editor
    ' for example
    ' Flasher1.Image = "fr"
    AllFlashOff()
End Sub

Sub AllFlashOff
    Dim i
    For i = 0 to 200
        FlashState(i) = 0
    Next
End Sub

Sub SetFlash(nr, stat)
    FlashState(nr) = ABS(stat)
End Sub

' Flasher objects
' Uses own faster timer
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
            If FlashLevel(nr) > 250 Then
                FlashLevel(nr) = 250
                FlashState(nr) = -2 'completely on
            End if
            Object.opacity = FlashLevel(nr)
    End Select
End Sub

Sub Flashm(nr, object) 'multiple flashers, it doesn't change the flashstate
    Select Case FlashState(nr)
        Case 0         'off
            Object.opacity = FlashLevel(nr)
        Case 1         ' on
            Object.opacity = FlashLevel(nr)
    End Select
End Sub

Sub FlashVal(nr, object, value)
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
            If FlashLevel(nr) > value Then
                FlashLevel(nr) = value
                FlashState(nr) = -2 'completely on
            End if
            Object.opacity = FlashLevel(nr)
    End Select
End Sub

Sub FlashValm(nr, object, value) 'multiple flashers, it doesn't change the flashstate
    Select Case FlashState(nr)
        Case 0 'off
            Object.opacity = FlashLevel(nr)
        Case 1 ' on
            Object.opacity = FlashLevel(nr)
    End Select
End Sub

Sub FadeLn(nr, Light, a, b, c, d)
    Select Case FadingLevel(nr)
        Case 2:Light.Offimage = d:FadingLevel(nr) = 0 'Off
        Case 3:Light.Offimage = c:FadingLevel(nr) = 2 'fading...
        Case 4:Light.Offimage = b:FadingLevel(nr) = 3 'fading...
        Case 5:Light.Offimage = a:FadingLevel(nr) = 1 'ON
    End Select
End Sub

Sub FadeLnm(nr, Light, a, b, c, d)
    Select Case FadingLevel(nr)
        Case 2:Light.Offimage = d
        Case 3:Light.Offimage = c
        Case 4:Light.Offimage = b
        Case 5:Light.Offimage = a
    End Select
End Sub

Sub LMapn(nr, Light, a, b, c, d)
    Select Case FadingLevel(nr)
        Case 2:Light.Onimage = d:FadingLevel(nr) = 0 'Off
        Case 3:Light.ONimage = c:FadingLevel(nr) = 2 'fading...
        Case 4:Light.ONimage = b:FadingLevel(nr) = 3 'fading...
        Case 5:Light.ONimage = a:FadingLevel(nr) = 1 'ON
    End Select
End Sub

Sub LMapnm(nr, Light, a, b, c, d)
    Select Case FadingLevel(nr)
        Case 2:Light.ONimage = d
        Case 3:Light.ONimage = c
        Case 4:Light.ONimage = b
        Case 5:Light.ONimage = a
    End Select
End Sub
' Walls

Sub FadeW(nr, a, b, c)
    Select Case FadingLevel(nr)
        Case 2:c.IsDropped = 1:FadingLevel(nr) = 0                 'Off
        Case 3:b.IsDropped = 1:c.IsDropped = 0:FadingLevel(nr) = 2 'fading...
        Case 4:a.IsDropped = 1:b.IsDropped = 0:FadingLevel(nr) = 3 'fading...
        Case 5:a.IsDropped = 0:FadingLevel(nr) = 1                 'ON
    End Select
End Sub

Sub FadeWm(nr, a, b, c)
    Select Case FadingLevel(nr)
        Case 2:c.IsDropped = 1
        Case 3:b.IsDropped = 1:c.IsDropped = 0
        Case 4:a.IsDropped = 1:b.IsDropped = 0
        Case 5:a.IsDropped = 0
    End Select
End Sub

Sub NFadeW(nr, a)
    Select Case FadingLevel(nr)
        Case 4:a.IsDropped = 1:FadingLevel(nr) = 0
        Case 5:a.IsDropped = 0:FadingLevel(nr) = 1
    End Select
End Sub

Sub NFadeWm(nr, a)
    Select Case FadingLevel(nr)
        Case 4:a.IsDropped = 1
        Case 5:a.IsDropped = 0
    End Select
End Sub

Sub NFadeWi(nr, a)
    Select Case FadingLevel(nr)
        Case 4:a.IsDropped = 0:FadingLevel(nr) = 0
        Case 5:a.IsDropped = 1:FadingLevel(nr) = 1
    End Select
End Sub

Sub NFadeWim(nr, a)
    Select Case FadingLevel(nr)
        Case 4:a.IsDropped = 0
        Case 5:a.IsDropped = 1
    End Select
End Sub

'Lights

Sub FadeL(nr, a, b)
    Select Case FadingLevel(nr)
        Case 2:b.state = 0:FadingLevel(nr) = 0
        Case 3:b.state = 1:FadingLevel(nr) = 2
        Case 4:a.state = 0:FadingLevel(nr) = 3
        Case 5:a.state = 1:FadingLevel(nr) = 1
    End Select
End Sub

Sub FadeLm(nr, a, b)
    Select Case FadingLevel(nr)
        Case 2:b.state = 0
        Case 3:b.state = 1
        Case 4:a.state = 0
        Case 5:a.state = 1
    End Select
End Sub

Sub LMap(nr, a, b, c) 'can be used with normal/olod style lights too
    Select Case FadingLevel(nr)
        Case 2:c.state = 0:FadingLevel(nr) = 0
        Case 3:b.state = 0:c.state = 1:FadingLevel(nr) = 2
        Case 4:a.state = 0:b.state = 1:FadingLevel(nr) = 3
        Case 5:b.state = 0:c.state = 0:a.state = 1:FadingLevel(nr) = 1
    End Select
End Sub

Sub LMapm(nr, a, b, c)
    Select Case FadingLevel(nr)
        Case 2:c.state = 0
        Case 3:b.state = 0:c.state = 1
        Case 4:a.state = 0:b.state = 1
        Case 5:b.state = 0:c.state = 0:a.state = 1
    End Select
End Sub

'Reels

Sub FadeR(nr, a)
    Select Case FadingLevel(nr)
        Case 2:a.SetValue 3:FadingLevel(nr) = 0
        Case 3:a.SetValue 2:FadingLevel(nr) = 2
        Case 4:a.SetValue 1:FadingLevel(nr) = 3
        Case 5:a.SetValue 1:FadingLevel(nr) = 1
    End Select
End Sub

Sub FadeRm(nr, a)
    Select Case FadingLevel(nr)
        Case 2:a.SetValue 3
        Case 3:a.SetValue 2
        Case 4:a.SetValue 1
        Case 5:a.SetValue 1
    End Select
End Sub

'Texts

Sub NFadeT(nr, a, b)
    Select Case FadingLevel(nr)
        Case 4:a.Text = "":FadingLevel(nr) = 0
        Case 5:a.Text = b:FadingLevel(nr) = 1
    End Select
End Sub

Sub NFadeTm(nr, a, b)
    Select Case FadingLevel(nr)
        Case 4:a.Text = ""
        Case 5:a.Text = b
    End Select
End Sub

' Flash a light, not controlled by the rom

Sub FlashL(nr, a, b)
    Select Case FadingLevel(nr)
        Case 1:b.state = 0:FadingLevel(nr) = 0
        Case 2:b.state = 1:FadingLevel(nr) = 1
        Case 3:a.state = 0:FadingLevel(nr) = 2
        Case 4:a.state = 1:FadingLevel(nr) = 3
        Case 5:b.state = 1:FadingLevel(nr) = 4
    End Select
End Sub

' Light acting as a flash. C is the light number to be restored

Sub MFadeL(nr, a, b, c)
    Select Case FadingLevel(nr)
        Case 2:b.state = 0:FadingLevel(nr) = 0:SetLamp c, FadingLevel(c)
        Case 3:b.state = 1:FadingLevel(nr) = 2
        Case 4:a.state = 0:FadingLevel(nr) = 3
        Case 5:a.state = 1:FadingLevel(nr) = 1
    End Select
End Sub

Sub MFadeLm(nr, a, b, c)
    Select Case FadingLevel(nr)
        Case 2:b.state = 0:SetLamp c, FadingLevel(c)
        Case 3:b.state = 1
        Case 4:a.state = 0
        Case 5:a.state = 1
    End Select
End Sub

'Alpha Ramps used as fading lights
'ramp is the name of the ramp
'a,b,c,d are the images used for on...off
'r is the refresh light

Sub FadeAR(nr, ramp, a, b, c, d)
    Select Case FadingLevel(nr)
        Case 2:ramp.image = d:FadingLevel(nr) = 0 'Off
        Case 3:ramp.image = c:FadingLevel(nr) = 2 'fading...
        Case 4:ramp.image = b:FadingLevel(nr) = 3 'fading...
        Case 5:ramp.image = a:FadingLevel(nr) = 1 'ON
    End Select
End Sub

Sub FadeARm(nr, ramp, a, b, c, d)
    Select Case FadingLevel(nr)
        Case 2:ramp.image = d
        Case 3:ramp.image = c
        Case 4:ramp.image = b
        Case 5:ramp.image = a
    End Select
End Sub

Sub FlashFO(nr, ramp, a, b, c)                                   'used for reflections when there is no off ramp
    Select Case FadingLevel(nr)
        Case 2:ramp.IsVisible = 0:FadingLevel(nr) = 0                'Off
        Case 3:ramp.image = c:FadingLevel(nr) = 2                'fading...
        Case 4:ramp.image = b:FadingLevel(nr) = 3                'fading...
        Case 5:ramp.image = a:ramp.IsVisible = 1:FadingLevel(nr) = 1 'ON
    End Select
End Sub

Sub FlashAR(nr, ramp, a, b, c)                                   'used for reflections when there is no off ramp
    Select Case FadingLevel(nr)
        Case 2:ramp.alpha = 0:FadingLevel(nr) = 0                'Off
        Case 3:ramp.image = c:FadingLevel(nr) = 2                'fading...
        Case 4:ramp.image = b:FadingLevel(nr) = 3                'fading...
        Case 5:ramp.image = a:ramp.alpha = 1:FadingLevel(nr) = 1 'ON
    End Select
End Sub

Sub FlashARm(nr, ramp, a, b, c)
    Select Case FadingLevel(nr)
        Case 2:ramp.alpha = 0
        Case 3:ramp.image = c
        Case 4:ramp.image = b
        Case 5:ramp.image = a:ramp.alpha = 1
    End Select
End Sub

Sub NFadeAR(nr, ramp, a, b)
    Select Case FadingLevel(nr)
        Case 4:ramp.image = b:FadingLevel(nr) = 0 'off
        Case 5:ramp.image = a:FadingLevel(nr) = 1 'on
    End Select
End Sub

Sub NFadeARm(nr, ramp, a, b)
    Select Case FadingLevel(nr)
        Case 4:ramp.image = b
        Case 5:ramp.image = a
    End Select
End Sub

Sub MNFadeAR(nr, ramp, a, b, c)
    Select Case FadingLevel(nr)
        Case 4:ramp.image = b:FadingLevel(nr) = 0:SetLamp c, FadingLevel(c) 'off
        Case 5:ramp.image = a:FadingLevel(nr) = 1                           'on
    End Select
End Sub

Sub MNFadeARm(nr, ramp, a, b, c)
    Select Case FadingLevel(nr)
        Case 4:ramp.image = b:SetLamp c, FadingLevel(c) 'off
        Case 5:ramp.image = a                           'on
    End Select
End Sub

' Flashers using PRIMITIVES
' pri is the name of the primitive
' a,b,c,d are the images used for on...off

Sub FadePri(nr, pri, a, b, c, d)
    Select Case FadingLevel(nr)
        Case 2:pri.image = d:FadingLevel(nr) = 0 'Off
        Case 3:pri.image = c:FadingLevel(nr) = 2 'fading...
        Case 4:pri.image = b:FadingLevel(nr) = 3 'fading...
        Case 5:pri.image = a:FadingLevel(nr) = 1 'ON
    End Select
End Sub

Sub FadePri3m(nr, pri, group)
    Select Case FadingLevel(nr)
        Case 3:pri.image = group(2)
        Case 4:pri.image = group(1)
        Case 5:pri.image = group(0)
    End Select
End Sub

Sub FadePri3(nr, pri, group)
    Select Case FadingLevel(nr)
        Case 3:pri.image = group(2):FadingLevel(nr) = 0 'Off
        Case 4:pri.image = group(1):FadingLevel(nr) = 3 'fading...
        Case 5:pri.image = group(0):FadingLevel(nr) = 1 'ON
    End Select
End Sub

Sub FadePri4m(nr, pri, group)
    Select Case FadingLevel(nr)
    Case 2:pri.image = group(3) 'Off
        Case 3:pri.image = group(2) 'Fading...
        Case 4:pri.image = group(1) 'Fading...
        Case 5:pri.image = group(0) 'ON
    End Select
End Sub

Sub FadePri4(nr, pri, group)
    Select Case FadingLevel(nr)
        Case 2:pri.image = group(3):FadingLevel(nr) = 0 'Off
        Case 3:pri.image = group(2):FadingLevel(nr) = 2 'Fading...
        Case 4:pri.image = group(1):FadingLevel(nr) = 3 'Fading...
        Case 5:pri.image = group(0):FadingLevel(nr) = 1 'ON
    End Select
End Sub

Sub FadePriC(nr, pri, a, b, c, d)
    Select Case FadingLevel(nr)
        Case 2:For each xx in pri:xx.image = d:Next:FadingLevel(nr) = 0 'Off
        Case 3:For each xx in pri:xx.image = c:Next:FadingLevel(nr) = 2 'fading...
        Case 4:For each xx in pri:xx.image = b:Next:FadingLevel(nr) = 3 'fading...
        Case 5:For each xx in pri:xx.image = a:Next:FadingLevel(nr) = 1 'ON
    End Select
End Sub

Sub FadePrih(nr, pri, a, b, c, d)
    Select Case FadingLevel(nr)
        Case 2:pri.image = d:SetFlash nr, 0:FadingLevel(nr) = 0 'Off
        Case 3:pri.image = c:FadingLevel(nr) = 2 'fading...
        Case 4:pri.image = b:FadingLevel(nr) = 3 'fading...
        Case 5:pri.image = a:SetFlash nr, 1:FadingLevel(nr) = 1 'ON
    End Select
End Sub

Sub FadePrim(nr, pri, a, b, c, d)
    Select Case FadingLevel(nr)
        Case 2:pri.image = d
        Case 3:pri.image = c
        Case 4:pri.image = b
        Case 5:pri.image = a
    End Select
End Sub

Sub NFadePri(nr, pri, a, b)
    Select Case FadingLevel(nr)
        Case 4:pri.image = b:FadingLevel(nr) = 0 'off
        Case 5:pri.image = a:FadingLevel(nr) = 1 'on
    End Select
End Sub

Sub NFadePrim(nr, pri, a, b)
    Select Case FadingLevel(nr)
        Case 4:pri.image = b
        Case 5:pri.image = a
    End Select
End Sub

'Fade a collection of lights

Sub FadeLCo(nr, a, b) 'fading collection of lights
    Dim obj
    Select Case FadingLevel(nr)
        Case 2:vpmSolToggleObj b, Nothing, 0, 0:FadingLevel(nr) = 0
        Case 3:vpmSolToggleObj b, Nothing, 0, 1:FadingLevel(nr) = 2
        Case 4:vpmSolToggleObj a, Nothing, 0, 0:FadingLevel(nr) = 3
        Case 5:vpmSolToggleObj a, Nothing, 0, 1:FadingLevel(nr) = 1
    End Select
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

' The sound is played using the VOL, PAN and PITCH functions, so the volume and pitch of the sound
' will change according to the ball speed, and the PAN function will change the stereo position according
' to the position of the ball on the table.


'**************************************
' Explanation of the collision routine
'**************************************

' The collision is built in VP.
' You only need to add a Sub OnBallBallCollision(ball1, ball2, velocity) and when two balls collide they
' will call this routine. What you add in the sub is up to you. As an example is a simple PlaysoundAtVol with volume and paning
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

Sub LeftFlipper2_Collide(parm)
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
    Case 1 : PlaySound "flip_hit_1", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 2 : PlaySound "flip_hit_2", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 3 : PlaySound "flip_hit_3", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End Select
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

Sub PlaySoundAtVol(sound, tableobj, Volum)
  PlaySound sound, 1, Volum, Pan(tableobj), 0,0,0, 1, AudioFade(tableobj)
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

Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "MGOW" is the name of the table
  Dim tmp
  tmp = tableobj.y * 2 / MGOW.height-1
  If tmp > 0 Then
    AudioFade = Csng(tmp ^10)
  Else
    AudioFade = Csng(-((- tmp) ^10) )
  End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "MGOW" is the name of the table
  Dim tmp
  tmp = tableobj.x * 2 / MGOW.width-1
  If tmp > 0 Then
    AudioPan = Csng(tmp ^10)
  Else
    AudioPan = Csng(-((- tmp) ^10) )
  End If
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "MGOW" is the name of the table
    Dim tmp
    tmp = ball.x * 2 / MGOW.width-1
    If tmp > 0 Then
        Pan = Csng(tmp ^10)
    Else
        Pan = Csng(-((- tmp) ^10) )
    End If
End Function

Function AudioFade(ball) ' Can this be together with the above function ?
  Dim tmp
  tmp = ball.y * 2 / MGOW.height-1
  If tmp > 0 Then
    AudioFade = Csng(tmp ^10)
  Else
    AudioFade = Csng(-((- tmp) ^10) )
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
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / (VolDiv/VolCol), Pan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub

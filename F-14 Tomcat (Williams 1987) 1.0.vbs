'**********************************************************
'    ____   _______   __________  __  ___________ ______
'   / __/__<  / / /  /_  __/ __ \/  |/  / ___/ _ /_  __/
'  / _//___/ /_  _/   / / / /_/ / /|_/ / /__/ __ |/ /
' /_/     /_/ /_/    /_/  \____/_/  /_/\___/_/ |_/_/
'
'**********************************************************
'
' Williams 1987, F-14 Tomcat
'
' Concept by:   Steve Ritchie
' Design by:  Steve Ritchie
' Art by:   Doug Watson
' Mechanics by:   Craig Fitpold
' Music by:   Steve Ritchie, Chris Granner
' Sound by:   Bill Parod, Chris Granner
' Software by:  Eugene Jarvis, Ed Boon
' Notes:  This game has two 7-digit alphanumeric score displays and two 7-digit numeric-only score displays.
'
' Credits:
'
' flupper1 - plastic ramp, wire ramps,
'            playfield shadows, primitive flippers
'
' assassin32 - script, playfield scan, table template
'
' ganjafarmer - plastics redraw, lighting, physics
'
'
' Thanks to: wrd1972, vogliadicane, senseless, Nemo, bjschneider93, bassgeige
'
' www.vpforums.org
'
'**********************************************************
Option Explicit
Randomize

' Thalamus 2018-07-23
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Changed UseSolenoids=1 to 2
' Thalamus 2018-08-26 : Improved directional sounds

' !! NOTE : Table not verified yet !!

Const VolDiv = 2000    ' Lower number, louder ballrolling/collition sound
Const VolCol    = 3    ' Ball collition divider ( voldiv/volcol )

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

Dim ChooseBats, Sticker, EdBoon, BumperColor, CustomBackwall, CustomCards

Const BallSize = 50 'do not change

'**********************************************************
'
'                          OPTIONS
'
'**********************************************************


ChooseBats = 2 '0=default flipper, 1=prim flipper yellow/red, 2=prim flipper white/red, 3=prim flipper white/blue

Sticker = 0 'aplly sticker on top of that beatiful ramp created by flupper1 - 0=OFF, 1=ON

EdBoon = 1 ' Yagov kicker with Ed Boon voice (Scorpion from Mortal Kombat 2) - 0=OFF, 1=ON

BumperColor = 2 '1=blue , 2=red

CustomBackwall = 1 'enable/disable custom backwall graphics

CustomCards = 0 'enable/disable custom cards


'**********************************************************
'
'                       END OF OPTIONS
'
'**********************************************************

Const cGameName="f14_l1",UseSolenoids=2,UseLamps=0,UseGI=0,SSolenoidOn="SolOn",SSolenoidOff="SolOff", SCoin="coin"

LoadVPM "01560000", "S11.VBS", 3.26
Dim DesktopMode: DesktopMode = Table1.ShowDT

If DesktopMode = True Then 'Show Desktop components
  Ramp16.visible=1
  Ramp15.visible=1
  Primitive13.visible=1
  ref1.visible=0
  ref2.visible=0
  ref3.visible=0
Else
  Ramp16.visible=0
  Ramp15.visible=0
  Primitive13.visible=0
End if

If Sticker = 1 Then
Ramp17.visible=1
Ramp18.visible=0
Else
Ramp17.visible=0
Ramp18.visible=1
End if

If BumperColor = 1 Then
      GI_Bumper_1.color = RGB(0,0,255)
      GI_Bumper_1.colorfull = RGB(0,0,255)

  Else
      Bumper.image = "bumper_red"
      GI_Bumper_1.color = RGB(255,0,0)
      GI_Bumper_1.colorfull = RGB(255,0,0)
End If

If CustomBackwall = 1 Then
  Wall21.sidevisible=1
Else
  Wall21.sidevisible=0
End If

If CustomCards = 1 Then
  Primitive68.image = "apron_custom"
End If

'*************************************************************
'Solenoid Call backs
'**********************************************************************************************************
SolCallback(1) = "bsTrough.SolIn"
SolCallback(2) = "bsTrough.SolOut"
SolCallback(3) = "SolPopper"
SolCallback(5) = "bsRC.SolOut"
SolCallback(6) = "vpmSolSound SoundFX(""Knocker"",DOFKnocker),"
SolCallback(7) = "bsRE.SolOut"
SolCallback(10) = "bsLC.SolOut"
SolCallback(12) = "bsKB.SolOut"
SolCallback(13) = "SolKickback"
Solcallback(14) = "PFGI"
SolCallback(16) = "SolRotateBeacons"
SolCallback(21) = "SolDiverter1"
SolCallback(22) = "SolDiverter2"

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
     If Enabled Then
         PlaySoundAtVol SoundFX("fx_Flipperup",DOFContactors), LeftFlipper, VolFlip:LeftFlipper.RotateToEnd:LeftFlipper1.RotateToEnd
     Else
         PlaySoundAtVol SoundFX("fx_Flipperdown",DOFContactors), LeftFlipper, VolFlip:LeftFlipper.RotateToStart:LeftFlipper1.RotateToStart
     End If
  End Sub

Sub SolRFlipper(Enabled)
     If Enabled Then
         PlaySoundAtVol SoundFX("fx_Flipperup",DOFContactors), RightFlipper, VolFlip:RightFlipper.RotateToEnd:RightFlipper1.RotateToEnd
     Else
         PlaySoundAtVol SoundFX("fx_Flipperdown",DOFContactors), RightFlipper, VolFlip:RightFlipper.RotateToStart:RightFlipper1.RotateToStart
     End If
End Sub

SolCallback(9) = "Flasher9" 'top blue
SolCallback(15) = "Flasher15" 'top red
SolCallback(25) = "Flasher25" 'yagov white
SolCallback(26) = "Flasher26" 'left lane
SolCallback(27) = "Flasher27" 'right lane
SolCallback(28) = "Flasher28" 'bottom red
SolCallback(29) = "Flasher29" 'bottom white
SolCallback(30) = "Flasher30" 'bottom blue
SolCallback(31) = "Flasher31" 'top white and red
SolCallback(32) = "SetLamp 132," 'radar flashers

Sub Flasher9(value)
  SetLamp 109, value
  If value Then
    FlasherDome9.disablelighting = 1
    FlasherDome9.image = "dome3_blue_lit"
    'FlasherDome9.material = "Flasherlit"
    PlaySoundAtVol "fx_relay", FlasherDome9, 1
  Else
    FlasherDome9.disablelighting = 0
    FlasherDome9.image = "dome3_blue"
    'FlasherDome9.material = "Flasherunlit"
  End If
End Sub

Sub Flasher15(value)
  SetLamp 115, value
  If value Then
    Primitive2.image="dome3_red_lit"
    'Primitive2.material="Flasherlit"
    Primitive2.disablelighting=1
    PlaySoundAtVol "fx_relay", Primitive2, 1
  Else
    Primitive2.image="dome3_red"
    'Primitive2.material="Flasherunlit"
    Primitive2.disablelighting=0
  End If
End Sub

Sub Flasher25(value)
  SetLamp 125, value
  if value Then
    Primitive4.image="dome3_clear_lit"
    'Primitive4.material="Flasherlit"
    Primitive4.disablelighting=1
    PlaySoundAtVol "fx_relay", Primitive4, 1
    Dome125.image="dome2_0_red_lit"
    Dome125.disablelighting=1
    topred.state=1 'backwall reflection
    Wall21.sideimage = "backdrop_lit"
    Wall21.disablelighting=1
  Else
    Primitive4.image="dome3_clear"
    'Primitive4.material="Flasherunlit"
    Primitive4.disablelighting=0
    Dome125.image="dome2_0_red_dark"
    Dome125.disablelighting=0
    topred.state=0 'backwall reflection
    Wall21.sideimage = "backdrop"
    Wall21.disablelighting=0
  End If
End Sub

Sub Flasher26(value)
  SetLamp 126, value
  if value Then
    Dome126.image="dome2_0_red_lit"
    Dome126.disablelighting=1
    topred.state=1 'backwall reflection
    'Wall21.sideimage = "backdrop_lit"
  Else
    Dome126.image="dome2_0_red_dark"
    Dome126.disablelighting=0
    topred.state=0 'backwall reflection
    'Wall21.sideimage = "backdrop"
  End If
End Sub

Sub Flasher27(value)
  SetLamp 127, value
  if value Then
    Dome127.image="dome2_0_red_lit"
    Dome127.disablelighting=1
    topred.state=1 'backwall reflection
    Wall21.sideimage = "backdrop_lit"
    Wall21.disablelighting=1
  Else
    Dome127.image="dome2_0_red_dark"
    Dome127.disablelighting=0
    topred.state=0 'backwall reflection
    Wall21.sideimage = "backdrop"
    Wall21.disablelighting=0
  End If
End Sub

Sub Flasher28(value)
  SetLamp 128, value
  if value Then
    Primitive16.image="dome3_red_lit"
    'Primitive16.material="Flasherlit"
    Primitive16.disablelighting=1
    PlaySoundAtVol "fx_relay", Primitive6, 1
    Dome128.image="dome2_0_red_lit"
    Dome128.disablelighting=1
    topred.state=1 'backwall reflection
    'Wall21.sideimage = "backdrop_lit"
  Else
    Primitive16.image="dome3_red"
    'Primitive16.material="Flasherunlit"
    Primitive16.disablelighting=0
    Dome128.image="dome2_0_red_dark"
    Dome128.disablelighting=0
    topred.state=0 'backwall reflection
    'Wall21.sideimage = "backdrop"
  End If
End Sub

Sub Flasher29(value)
  SetLamp 129, value
  if value Then
    Primitive5.image="dome3_clear_lit"
    'Primitive5.material="Flasherlit"
    Primitive5.disablelighting=1
    PlaySoundAtVol "fx_relay", Primitive5, 1
    Dome129.image="dome2_0_red_lit"
    Dome129.disablelighting=1
    topred.state=1 'backwall reflection
    Wall21.sideimage = "backdrop_lit"
    Wall21.disablelighting=1
  Else
    Primitive5.image="dome3_clear"
    'Primitive5.material="Flasherunlit"
    Primitive5.disablelighting=0
    Dome129.image="dome2_0_red_dark"
    Dome129.disablelighting=0
    topred.state=0 'backwall reflection
    Wall21.sideimage = "backdrop"
    Wall21.disablelighting=0
  End If
End Sub

Sub Flasher30(value)
  SetLamp 130, value
  if value Then
    Primitive6.image="dome3_blue_lit"
    'Primitive6.material="Flasherlit"
    Primitive6.disablelighting=1
    PlaySoundAtVol "fx_relay", Primitive6, 1
    Dome130.image="dome2_0_red_lit"
    Dome130.disablelighting=1
    topred.state=1 'backwall reflection
    'Wall21.sideimage = "backdrop_lit"
  Else
    Primitive6.image="dome3_blue"
    'Primitive6.material="Flasherunlit"
    Primitive6.disablelighting=0
    Dome130.image="dome2_0_red_dark"
    Dome130.disablelighting=0
    topred.state=0 'backwall reflection
    'Wall21.sideimage = "backdrop"
  End If
End Sub

Sub Flasher31(value)
  SetLamp 131, value
  if value Then
    Primitive3.image="dome3_clear_lit"
    'Primitive3.material="Flasherlit"
    Primitive3.disablelighting=1
    Primitive1.image="dome3_red_lit"
    'Primitive1.material="Flasherlit"
    Primitive1.disablelighting=1
    PlaySoundAtVol "fx_relay", Primitive3, 1
  Else
    Primitive3.image="dome3_clear"
    'Primitive3.material="Flasherunlit"
    Primitive3.disablelighting=0
    Primitive1.image="dome3_red"
    'Primitive1.material="Flasherunlit"
    Primitive1.disablelighting=0
  End If
End Sub

'**********************************************************************************************************

'Solenoid Controlled toys
'**********************************************************************************************************

'Playfield GI
Sub PFGI(Enabled)

  If Enabled Then
    dim xx
    For each xx in GI:xx.State = 1: Next
        PlaySound "fx_relay"
    ref1.visible=1
    ref2.visible=1
    ref3.visible=1
  Else
    For each xx in GI:xx.State = 0: Next
        PlaySound "fx_relay"
    ref1.visible=0
    ref2.visible=0
    ref3.visible=0
  End If
End Sub


Dim popperBall, popperZpos

Sub SolPopper(Enabled)
    If Enabled Then
        If bsBP.Balls Then
            Set popperBall = sw24a.Createball
            popperBall.Z = 0
            popperZpos = 0
      PlaySoundAtVol SoundFX("Popper",DOFContactors), sw24, VolKick
            sw24a.TimerInterval = 2
            sw24a.TimerEnabled = 1
        End If
    End If
End Sub

Sub sw24a_Timer
    popperBall.Z = popperZpos
    popperZpos = popperZpos + 10
    If popperZpos> 210 Then
        sw24a.TimerEnabled = 0
        sw24a.DestroyBall
        bsBP.ExitSol_On
    End If
End Sub


Sub SolKickBack(enabled)
    If enabled Then
        KickBack.Fire
    PlaySoundAtVol SoundFX("Popper",DOFContactors), KickBack, VolKick
    Else
        KickBack.PullBack
    End If
End Sub

Sub SolDiverter1(enabled)
    If enabled Then
        Diverter1.RotateToEnd
    PlaySoundAtVol SoundFX("Popper_Ball",DOFContactors), Diverter1, 1
    PlaySoundAtVol "fx_relay", Diverter1, 1
    Else
        Diverter1.RotateToStart
    PlaySoundAtVol SoundFX("Popper_Ball",DOFContactors), Diverter1, 1
    PlaySoundAtVol "fx_relay", Diverter1, 1
    End If
End Sub

Sub SolDiverter2(enabled)
    If enabled Then
        Diverter2.RotateToEnd
    PlaySoundAtVol SoundFX("Popper_Ball",DOFContactors), Diverter2, 1
    PlaySoundAtVol "fx_relay", Diverter2, 1
    Else
        Diverter2.RotateToStart
    PlaySoundAtVol SoundFX("Popper_Ball",DOFContactors), Diverter2, 1
    PlaySoundAtVol "fx_relay", Diverter2, 1
    End If
End Sub


Sub SolRotateBeacons(Enabled)
    If enabled then

    Else

    End If
End Sub



'**********************************************************************************************************

'Initiate Table
'**********************************************************************************************************

dim bsTrough, bsRE, bsLC, bsRC, bsBP, bsKB

Sub Table1_Init
  vpmInit Me
  On Error Resume Next
    With Controller
    .GameName = cGameName
    If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
    .SplashInfoLine = "F-14 Tomcat (Williams 1987)"&chr(13)&"You Suck"
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

    PinMAMETimer.Interval = PinMAMEInterval
    PinMAMETimer.Enabled = 1

    vpmNudge.TiltSwitch = swTilt
    vpmNudge.Sensitivity = 5
    vpmNudge.TiltObj = Array(Bumper1, LeftSlingshot, RightSlingshot)

  'Change flippers

  ChangeBats(ChooseBats)

    ' Trough
    Set bsTrough = New cvpmBallStack
    bsTrough.InitSw 10, 14, 13, 12, 11, 0, 0, 0
    bsTrough.InitKick BallRelease, 135, 6
    bsTrough.InitExitSnd SoundFX("ballrelease",DOFContactors), SoundFX("Solenoid",DOFContactors)
    bsTrough.Balls = 4

    ' Right Eject
    Set bsRE = New cvpmBallStack
    bsRE.InitSaucer sw21, 21, 0, 38
    bsRE.InitExitSnd SoundFX("Popper",DOFContactors), SoundFX("Solenoid",DOFContactors)

    ' Left Center Eject
    Set bsLC = New cvpmBallStack
    bsLC.InitSaucer sw22, 22, 335, 45
    bsLC.InitExitSnd SoundFX("Popper",DOFContactors), SoundFX("Solenoid",DOFContactors)

    ' Right Center Eject
    Set bsRC = New cvpmBallStack
    bsRC.InitSaucer sw23, 23, 27, 45
    bsRC.InitExitSnd SoundFX("Popper",DOFContactors), SoundFX("Solenoid",DOFContactors)
    bsRC.KickAngleVar = 4
    bsRC.KickForceVar = 1

    ' Ball Popper
    Set bsBP = New cvpmBallStack
    bsBP.InitSaucer sw24, 24, 0, 0
    bsBP.InitKick sw24a, 270, 28
    bsBP.InitExitSnd SoundFX("Popper",DOFContactors), SoundFX("Solenoid",DOFContactors)

    ' Line of Death Kickback
    Set bsKB = New cvpmBallStack
    bsKB.InitSaucer sw55, 55, 178, 55
    bsKB.InitExitSnd SoundFX("Popper",DOFContactors), SoundFX("Solenoid",DOFContactors)
    bsKB.KickAngleVar = 1
    bsKB.KickForceVar = 5

    ' Init ramp diverter
    Diverter1.RotatetoStart
    Diverter2.RotatetoStart

    ' Init Rescue Kickback
    KickBack.Pullback

End Sub

'Change flippers

Sub ChangeBats(Bats)
  Select Case Bats
    Case 0
      batleft.visible = 0 : batright.visible = 0 : batleft1.visible = 0 : batright1.visible = 0 : LeftFlipper.visible = 1 : RightFlipper.visible = 1 : LeftFlipper1.visible = 1 : RightFlipper1.visible = 1
      batleftshadow.visible = 0 : batrightshadow.visible = 0 : batleftshadow1.visible = 0 : batrightshadow1.visible = 0 : GraphicsTimer.enabled = False
    Case 1
      batleft.visible = 1 : batright.visible = 1 : batleft1.visible = 1 : batright1.visible = 1 : LeftFlipper.visible = 0 : RightFlipper.visible = 0 : LeftFlipper1.visible = 0 : RightFlipper1.visible = 0
      batleftshadow.visible = 1 : batrightshadow.visible = 1 : batleftshadow1.visible = 1 : batrightshadow1.visible = 1 : GraphicsTimer.enabled = True
    Case 2
      batleft.visible = 1 : batright.visible = 1 : batleft1.visible = 1 : batright1.visible = 1 : LeftFlipper.visible = 0 : RightFlipper.visible = 0 : LeftFlipper1.visible = 0 : RightFlipper1.visible = 0
      batleftshadow.visible = 1 : batrightshadow.visible = 1 : batleftshadow1.visible = 1 : batrightshadow1.visible = 1 : GraphicsTimer.enabled = True
      batleft.image = "flipper_white_red" : batright.image = "flipper_white_red" : batleft1.image = "flipper_white_red" : batright1.image = "flipper_white_red"
    Case 3
      batleft.visible = 1 : batright.visible = 1 : batleft1.visible = 1 : batright1.visible = 1 : LeftFlipper.visible = 0 : RightFlipper.visible = 0 : LeftFlipper1.visible = 0 : RightFlipper1.visible = 0
      batleftshadow.visible = 1 : batrightshadow.visible = 1 : batleftshadow1.visible = 1 : batrightshadow1.visible = 1 : GraphicsTimer.enabled = True
      batleft.image = "flipper_white_blue" : batright.image = "flipper_white_blue" : batleft1.image = "flipper_white_blue" : batright1.image = "flipper_white_blue"
  End Select
End Sub

'**********************************************************************************************************
'Plunger code
'**********************************************************************************************************

Sub Table1_KeyDown(ByVal KeyCode)
  If KeyDownHandler(keycode) Then Exit Sub
  If keycode = PlungerKey Then Plunger.Pullback:playsoundAtVol "plungerpull", Plunger, 1
    If keycode = RightFlipperKey Then Controller.Switch(63) = 1
    If keycode = LeftFlipperKey Then Controller.Switch(15) = 1
End Sub

Sub Table1_KeyUp(ByVal KeyCode)
  If KeyUpHandler(keycode) Then Exit Sub
  If keycode = PlungerKey Then Plunger.Fire:PlaySoundAtVol "plunger", Plunger, 1
    If keycode = RightFlipperKey Then Controller.Switch(63) = 0
    If keycode = LeftFlipperKey Then Controller.Switch(15) = 0
End Sub

'**********************************************************************************************************

 ' Drain and Kickers
Sub Drain_Hit:bsTrough.addball me : playsoundAtVol "drain", Drain, 1 : End Sub
Sub sw21_Hit:bsRE.AddBall Me : playsoundAtVol "popper_ball", sw21, 1 : End Sub
Sub sw22_Hit:bsLC.AddBall Me : playsoundAtVol "popper_ball", sw22, 1 : End Sub
Sub sw23_Hit:bsRC.AddBall Me : playsoundAtVol "popper_ball", sw23, 1 : End Sub
Sub sw24_Hit:bsBP.AddBall Me : playsoundAtVol "popper_ball", sw24, 1 : End Sub

'YAGOV KICKER

Sub sw55_Hit
  bsKB.AddBall Me
  playsoundAtVol "popper_ball", sw55, VolKick

  If EdBoon = 1 Then
    Select Case Int(Rnd*2)+1
      Case 1 : playsoundAtVol "scorpion", sw55, 1
      Case 2 : playsoundAtVol "scorpion2", sw55, 1
    End Select
  End If

End Sub


'Bumpers
Sub Bumper1_Hit : vpmTimer.PulseSw(28) : playsoundAtVol SoundFX("fx_bumper",DOFContactors), Bumper1, VolBump: End Sub

' Stand Up Targets
Sub sw25_Hit:vpmTimer.PulseSw(25):End Sub
Sub sw26_Hit:vpmTimer.PulseSw(26):End Sub
Sub sw33_Hit:vpmTimer.PulseSw(33):End Sub
Sub sw34_Hit:vpmTimer.PulseSw(34):End Sub
Sub sw35_Hit:vpmTimer.PulseSw(35):End Sub
Sub sw36_Hit:vpmTimer.PulseSw(36):End Sub
Sub sw37_Hit:vpmTimer.PulseSw(37):End Sub
Sub sw38_Hit:vpmTimer.PulseSw(38):End Sub
Sub sw41_Hit:vpmTimer.PulseSw(41):End Sub
Sub sw42_Hit:vpmTimer.PulseSw(42):End Sub
Sub sw43_Hit:vpmTimer.PulseSw(43):End Sub
Sub sw44_Hit:vpmTimer.PulseSw(44):End Sub
Sub sw45_Hit:vpmTimer.PulseSw(45):End Sub
Sub sw46_Hit:vpmTimer.PulseSw(46):End Sub
Sub sw49_Hit:vpmTimer.PulseSw(49):End Sub
Sub sw50_Hit:vpmTimer.PulseSw(50):End Sub
Sub sw51_Hit:vpmTimer.PulseSw(51):End Sub
Sub sw52_Hit:vpmTimer.PulseSw(52):End Sub
Sub sw53_Hit:vpmTimer.PulseSw(53):End Sub
Sub sw54_Hit:vpmTimer.PulseSw(54):End Sub

' Wire Triggers
Sub sw16_Hit:Controller.Switch(16) = 1 : playsoundAtVol "rollover", sw16, VolRol : End Sub
Sub sw16_Unhit:Controller.Switch(16) = 0:End Sub
Sub sw59_Hit:Controller.Switch(59) = 1 : playsoundAtVol "rollover", sw59, VolRol : End Sub
Sub sw59_UnHit:Controller.Switch(59) = 0:End Sub
Sub sw60_Hit:Controller.Switch(60) = 1 : playsoundAtVol "rollover", sw60, VolRol : End Sub
Sub sw60_UnHit:Controller.Switch(60) = 0:End Sub
Sub sw61_Hit:Controller.Switch(61) = 1 : playsoundAtVol "rollover", sw61, VolRol : End Sub
Sub sw61_UnHit:Controller.Switch(61) = 0:End Sub
Sub sw62_Hit:Controller.Switch(62) = 1 : playsoundAtVol "rollover", sw62, VolRol : End Sub
Sub sw62_UnHit:Controller.Switch(62) = 0:End Sub

'Ramp Triggers
Sub sw20_Hit:Controller.Switch(20) = 1 : playsoundAtVol "fx_metalrolling", sw20, VolMetal : End Sub
Sub sw20_Unhit:Controller.Switch(20) = 0:End Sub
Sub sw30_Hit:Controller.Switch(30) = 1 : playsoundAtVol "fx_metalrolling", sw30, VolMetal : End Sub
Sub sw30_Unhit:Controller.Switch(30) = 0:End Sub
Sub sw31_Hit:Controller.Switch(31) = 1 : playsoundAtVol "fx_rrenter", sw31, VolMetal : End Sub
Sub sw31_Unhit:Controller.Switch(31) = 0:End Sub
Sub sw32_Hit:Controller.Switch(32) = 1 : playsoundAtVol "fx_rrenter", sw32, VolMetal : End Sub
Sub sw32_Unhit:Controller.Switch(32) = 0:End Sub

Sub Trigger1_hit: playsoundAtVol "fx_rr5", Trigger1, 1 : End Sub
Sub Trigger2_hit: playsoundAtVol "fx_rr6", Trigger2, 1 : End Sub
Sub Trigger3_hit: playsoundAtVol "fx_rr7", Trigger3, 1 : End Sub
Sub Trigger4_hit: playsoundAtVol "fx_rr5", Trigger4, 1 : End Sub

'Gate Triggers
Sub sw47_Hit:vpmTimer.PulseSw(47):End Sub
Sub sw56_Hit:vpmTimer.PulseSw(56):End Sub

'Spinners
Sub sw48_Spin:vpmTimer.PulseSw 48 : playsoundAtVol "fx_spinner", sw48, VolSpin : End Sub


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
    NFadeL 8, l8
    NFadeL 9, l9
    NFadeL 10, l10
    NFadeL 11, l11
    NFadeL 12, l12
    NFadeL 13, l13
    NFadeL 14, l14
    NFadeL 15, l15
    NFadeL 16, l16
    NFadeL 17, l17
    NFadeL 18, l18
    NFadeL 19, l19
    NFadeL 20, l20
    NFadeL 21, l21
    NFadeL 22, l22
    NFadeL 23, l23
    NFadeLm 24, l24a
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
    NFadeLm 39, l39a
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
    NFadeLm 57, l57a 'Left Ramp LEDs
    NFadeObj 57, l57, "bulbcover1_redOn", "bulbcover1_red"
    NFadeLm 58, l58a  'Left Ramp LEDs
    NFadeObj 58, l58, "bulbcover1_redOn", "bulbcover1_red"
    NFadeLm 59, l59a  'Left Ramp LEDs
    NFadeObj 59, l59, "bulbcover1_redOn", "bulbcover1_red"
    NFadeLm 60, l60a  'Left Ramp LEDs
    NFadeObj 60, l60, "bulbcover1_blueOn", "bulbcover1_blue"
    NFadeLm 61, l61a  'Left Ramp LEDs
    NFadeObj 61, l61, "bulbcover1_blueOn", "bulbcover1_blue"
    NFadeLm 62, l62a  'Left Ramp LEDs
    NFadeObj 62, l62, "bulbcover1_blueOn", "bulbcover1_blue"
    NFadeL 63, l63
    NFadeL 64, l64

'Solenoid Controlled lamps
    NFadeLm 109, f109
    NFadeLm 109, f109a

    NFadeLm 115, f115
    NFadeLm 115, f115a

    NFadeLm 125, f125a
    NFadeLm 125, f125b
    NFadeLm 125, f125 'Backwall

    NFadeLm 126, f126a
    NFadeL 126, f126 'Backwall

    NFadeLm 127, f127a
    NFadeL 127, f127 'Backwall

    NFadeLm 128, f128a
    NFadeLm 128, f128b
    NFadeL  128, f128 'Backwall

    NFadeLm 129, f129a
    NFadeLm 129, f129b
    NFadeL  129, f129 'Backwall

    NFadeLm 130, f130a
    NFadeLm 130, f130b
    NFadeL  130, f130 'Backwall

    NFadeLm 131, f131
    NFadeLm 131, f131a
    NFadeLm 131, f131b
    NFadeLm 131, f131c

    NFadeLm 132, f132
  NFadeLm 132, f132b

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
 Dim Digits(28)
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

 ' 3rd Player
Digits(14) = Array(LED150,LED151,LED152,LED153,LED154,LED155,LED156)
Digits(15) = Array(LED160,LED161,LED162,LED163,LED164,LED165,LED166)
Digits(16) = Array(LED170,LED171,LED172,LED173,LED174,LED175,LED176)
Digits(17) = Array(LED180,LED181,LED182,LED183,LED184,LED185,LED186)
Digits(18) = Array(LED190,LED191,LED192,LED193,LED194,LED195,LED196)
Digits(19) = Array(LED200,LED201,LED202,LED203,LED204,LED205,LED206)
Digits(20) = Array(LED210,LED211,LED212,LED213,LED214,LED215,LED216)

' 4th Player
Digits(21) = Array(LED220,LED221,LED222,LED223,LED224,LED225,LED226)
Digits(22) = Array(LED230,LED231,LED232,LED233,LED234,LED235,LED236)
Digits(23) = Array(LED240,LED241,LED242,LED243,LED244,LED245,LED246)
Digits(24) = Array(LED250,LED251,LED252,LED253,LED254,LED255,LED256)
Digits(25) = Array(LED260,LED261,LED262,LED263,LED264,LED265,LED266)
Digits(26) = Array(LED270,LED271,LED272,LED273,LED274,LED275,LED276)
Digits(27) = Array(LED280,LED281,LED282,LED283,LED284,LED285,LED286)


 Sub DisplayTimer_Timer
    Dim ChgLED, ii, jj, num, chg, stat, obj, b, x
    ChgLED=Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
    If Not IsEmpty(ChgLED)Then
    If DesktopMode = True Then
       For ii=0 To UBound(chgLED)
          num=chgLED(ii, 0) : chg=chgLED(ii, 1) : stat=chgLED(ii, 2)
      if (num < 28) then
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


' *********************************************************************
' *********************************************************************

          'Start of VPX call back Functions

' *********************************************************************
' *********************************************************************

'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep, Cstep

Sub RightSlingShot_Slingshot
  vpmTimer.PulseSw 58
    PlaySoundAtVol SoundFX("fx_slingshot",DOFContactors), sling1, 1
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.TransZ = -20
    RStep = 0
    RightSlingShot.TimerEnabled = 1
  f127a.state=1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.TransZ = -10:
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.TransZ = 0:RightSlingShot.TimerEnabled = 0:f127a.state=0:PlaySoundAtVol "fx_relay", sling2, 1:
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
  vpmTimer.PulseSw 57
    PlaySoundAtVol SoundFX("fx_slingshot",DOFContactors), sling2, 1
    LSling.Visible = 0
    LSling1.Visible = 1
    sling2.TransZ = -20
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
  f126a.state=1
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.TransZ = -10:
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.TransZ = 0:LeftSlingShot.TimerEnabled = 0:f126a.state=0:PlaySoundAtVol "fx_relay", sling1, 1:
    End Select
    LStep = LStep + 1
End Sub

Sub TopRubberSling_Hit
  RubberTop.visible = 0
    RubberTop1.Visible = 0
    RubberTop2.Visible = 1
    CStep = 0
    TopRubberSling.TimerEnabled = 1
End Sub

Sub TopRubberSling_Timer
    Select Case CStep
        Case 3:RubberTop1.Visible = 0:RubberTop2.Visible = 1:
        Case 4:RubberTop2.Visible = 0:RubberTop.Visible = 1:TopRubberSling.TimerEnabled = 0:
    End Select
    CStep = CStep + 1
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
  PlaySound "pinhit_low", 0, Vol(ActiveBall)*VolPi, Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Targets_Hit (idx)
  PlaySound "target", 0, Vol(ActiveBall)*VolTarg, Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Metals_Thin_Hit (idx)
  PlaySound "metalhit_thin", 0, Vol(ActiveBall)*VolMetal, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Metals_Medium_Hit (idx)
  PlaySound "metalhit_medium", 0, Vol(ActiveBall)*VolMetal, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Metals2_Hit (idx)
  PlaySound "metalhit2", 0, Vol(ActiveBall)*VolMetal, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Gates_Hit (idx)
  PlaySound "gate4", 0, Vol(ActiveBall)*VolGates, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Spinner_Spin
  PlaySoundAtVol "fx_spinner", Spinner, 1
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
    Case 1 : PlaySound "flip_hit_1", 0, Vol(ActiveBall)*VolFlip, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 2 : PlaySound "flip_hit_2", 0, Vol(ActiveBall)*VolFlip, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 3 : PlaySound "flip_hit_3", 0, Vol(ActiveBall)*VolFlip, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End Select
End Sub

Sub GraphicsTimer_Timer()
    If ChooseBats > 0 Then
    ' *** move primitive bats ***
    batleft.objrotz = LeftFlipper.CurrentAngle + 1
    batleftshadow.objrotz = batleft.objrotz
    batright.objrotz = RightFlipper.CurrentAngle - 1
    batrightshadow.objrotz  = batright.objrotz
    batleft1.objrotz = LeftFlipper1.CurrentAngle + 1
    batleftshadow1.objrotz = batleft1.objrotz
    batright1.objrotz = RightFlipper1.CurrentAngle - 1
    batrightshadow1.objrotz  = batright1.objrotz
  End If
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

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "table1" is the name of the table
    Dim tmp
    tmp = ball.x * 2 / table1.width-1
    If tmp > 0 Then
        Pan = Csng(tmp ^10)
    Else
        Pan = Csng(-((- tmp) ^10) )
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

' Thalamus : Exit in a clean and proper way
Sub Table1_exit()
  Controller.Pause = False
  Controller.Stop
End Sub


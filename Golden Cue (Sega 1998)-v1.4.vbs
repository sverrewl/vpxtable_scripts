' *********************************************************************
' *********************************************************************
' **                    G O L D E N   C U E                          **
' **                                                                 **
' **                      Table Script v1.4                          **
' **                      2021, 2022 Felsir                          **
' **                                                                 **
' *********************************************************************
' *********************************************************************

' Thank you to all of the other virtual pinball content creators,
' I stand on the shoulders of giants!
'
' Special thank you to 32assassin with providing me of additional
' photo's of the playfield.

' v0.6  First public WIP release
' v0.7  Increased playfield resolution (thanks Timblo)
'   Updated droptarget textures
'   Added plunger cover texture
'   Added flasher mesh
' v0.8  Added more platics graphics
'   Tweaked flashers.
'   Changed the flippers
' v0.9  Added screws and pegs
'   Visual tweaks
'   Adjusted flashers
'   Chose to add billiard ball graphics, the photo quality of Kelly on the right hand side is not good enough.
' v1.0  Prepped for release- DT graphic added
' v1.1  Added Kelly substitute graphic on the righthand side plastic. (thanks Aldiode)
' v1.2  Fixed the Post-Save buttons (thanks Pinballuser)
'   Added DOF and FX for bumpers and autoplunger (thanks Pinballuser)
' v1.3  Updated playfield with better graphic of Kelly's face (thanks Timblo)
'   Fixed the wall behind the extra left flipper - ball could fly past it.
' v1.4  Fixed an geometry issue that sometimes caused the multiball to miscount the number of balls (thanks Aldiode)
'
'
' TODOs:
' Improve flasher lights (they work but are not yet flashing as I wish they would...
' Improve playfield plastics, graphics are hard to find.
' Add bolts, screws, posts for a more realistic appearance.
' Tweak values for physics
' Playtest more ...

Option Explicit
Randomize

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Const cGameName="goldcue",UseSolenoids=1,UseLamps=0,UseGI=0, SCoin="coin"

LoadVPM "01320000", "SEGA.VBS", 3.10

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
SolCallBack(1)      = "SolTrough"
SolCallBack(2)    = "SolAutofire"
SolCallback(3)    = "SetLamp 103,"  'Lower Left
SolCallback(4)    = "SolDiv"      '8-Ball Diverter
SolCallback(5)    = "SolLock"
SolCallback(6)    = "dtLL.SolDropUp"
SolCallback(7)    = "dtDrop.SolDropUp"
'SolCallback(9)   = "vpmSolSound ""jet3"","
'SolCallback(10)  = "vpmSolSound ""jet3"","
'SolCallback(11)  = "vpmSolSound ""jet3"","
'SolCallBack(12)  = "vpmSolSound ""sling"","
'SolCallBack(13)  = "vpmSolSound ""sling"","
SolCallBack(17)     = "SetLamp 117,"  'Upper Left x2
SolCallBack(18)     = "SetLamp 118,"  'Flash Uppper Right X3
SolCallBack(19)   = "SetLamp 119,"  'Flash Mid Right
SolCallBack(20)     = "SetLamp 120,"  'Flash Lower Right
SolCallBack(21)   = "SolLPost"    'Left Outlane
SolCallBack(22)   = "SolRPost"    'Right Outlane
SolCallBack(23)   = "SolPostUp"   'Up/Down Post
SolCallBack(25)   = "dtDrop.SolHit 1,"
SolCallBack(26)   = "dtDrop.SolHit 2,"
SolCallBack(27)   = "dtDrop.SolHit 3,"
SolCallBack(28)   = "dtDrop.SolHit 4,"
SolCallBack(29)   = "dtDrop.SolHit 5,"
SolCallBack(30)   = "dtDrop.SolHit 6,"
SolCallBack(31)   = "dtDrop.SolHit 7,"
SolCallback(32)   = "dtLL.SolHit 1,"

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
     If Enabled Then
         PlaySoundAtVol SoundFX("fx_Flipperup",DOFContactors), LeftFlipper, 1:LeftFlipper.RotateToEnd:LeftFlipper1.RotateToEnd
     Else
         PlaySoundAtVol SoundFX("fx_Flipperdown",DOFContactors), LeftFlipper, 1:LeftFlipper.RotateToStart:LeftFlipper1.RotateToStart
     End If
  End Sub

Sub SolRFlipper(Enabled)
     If Enabled Then
         PlaySoundAtVol SoundFX("fx_Flipperup",DOFContactors), RightFlipper, 1:RightFlipper.RotateToEnd
     Else
         PlaySoundAtVol SoundFX("fx_Flipperdown",DOFContactors), RightFlipper, 1:RightFlipper.RotateToStart
     End If
End Sub
'**********************************************************************************************************

'Solenoid Controlled toys
'**********************************************************************************************************

Sub SolTrough(Enabled)
  If Enabled Then
    bsTrough.ExitSol_On
    vpmTimer.PulseSw 15
  End If
End Sub

Sub solAutofire(Enabled)
  If Enabled Then
    PlungerIM.AutoFire
  End If
End Sub

Sub SolDiv(Enabled)
     If Enabled Then
    Divert.IsDropped = 1
        PlaySoundAtVol SoundFX("fx_Flipperup",DOFContactors), Divert, 1
     Else
    Divert.IsDropped = 0
        PlaySoundAtVol SoundFX("fx_Flipperdown",DOFContactors), Divert, 1
     End If
End Sub

Sub SolLock(Enabled)
  If Enabled Then
    LockBPost.IsDropped=1
        PlaySoundAtVol SoundFX("fx_Flipperup",DOFContactors), LockBPost, 1
  Else
    LockBPost.IsDropped=0
        PlaySoundAtVol SoundFX("fx_Flipperdown",DOFContactors), LockBPost, 1
  End If
End Sub

Sub SolLPost(Enabled)
  If Enabled Then
    LeftPost.IsDropped=0
        PlaySoundAtVol SoundFX("Popper",DOFContactors), LeftPost, 1
  Else
    LeftPost.IsDropped=1
        PlaySoundAtVol SoundFX("Popper",DOFContactors), LeftPost, 1
  End If
End Sub

Sub SolRPost(Enabled)
  If Enabled Then
    RightPost.IsDropped=0
        PlaySoundAtVol SoundFX("Popper",DOFContactors), RightPost, 1
  Else
    RightPost.IsDropped=1
        PlaySoundAtVol SoundFX("Popper",DOFContactors), RightPost, 1
  End If
End Sub

Sub SolPostUp(Enabled)
  If Enabled Then
    CenterPost.IsDropped=0
    CenterPostPrim.transY = 23
  ' L24.BulbHaloHeight = 22 ' Need to add don't know Lamp ID
    PlaysoundAtVol SoundFX("Centerpost_Up",DOFContactors), CenterPostPrim, 1
  Else
    CenterPost.IsDropped=1
    CenterPostPrim.transY = 0
  ' L24.BulbHaloHeight = 1 ' Need to add don't know Lamp ID
    PlaysoundAtVol SoundFX("Centerpost_Down",DOFContactors), CenterPostPrim, 1
  End If
End Sub

'Stern-Sega GI
set GICallback = GetRef("UpdateGI")
Sub UpdateGI(no, Enabled)
    SetLamp 101, Enabled  'dummy GI call for B2s server for an animated backglass
  If Enabled Then
    Dim xx
    For each xx in GI:xx.State = 1:Next
    PlaySound "fx_relay"
  Else
    For each xx in GI:xx.State = 0:Next
    PlaySound "fx_relay"
  End If
End Sub

'**********************************************************************************************************

'Initiate Table
'**********************************************************************************************************

Dim bsTrough, dtDrop, dtLL

Sub Table1_Init
  vpmInit Me
  On Error Resume Next
    With Controller
    .GameName = cGameName
    If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
    .SplashInfoLine = "Golden Cue Sega"&chr(13)&"Enjoy!"
    .HandleMechanics=0
    .HandleKeyboard=0
    .ShowDMDOnly=1
    .ShowFrame=0
    .ShowTitle=0
    .hidden = 0
        .Games(cGameName).Settings.Value("sound")=1
    '.PuPHide = 1
         On Error Resume Next
         .Run GetPlayerHWnd
         If Err Then MsgBox Err.Description
         On Error Goto 0
     End With
     On Error Goto 0

  PinMAMETimer.Interval=PinMAMEInterval
  PinMAMETimer.Enabled=1

  vpmNudge.TiltSwitch=56
  vpmNudge.Sensitivity=2
  vpmNudge.TiltObj=Array(Bumper1,Bumper2,Bumper3,LeftSlingshot,RightSlingshot)

    Set bsTrough = New cvpmBallStack
    bsTrough.InitSw 0,14,13,12,11,0,0,0
    bsTrough.InitKick BallRelease,90,3
    bsTrough.InitExitSnd SoundFX("ballrelease",DOFContactors), SoundFX("Solenoid",DOFContactors)
    bsTrough.Balls=4

'    Set bsTrough = New cvpmTrough
'    With bsTrough
'   .Size = 4
'   .InitSwitches Array(11, 12, 13, 14)
'   .InitExit BallRelease, 90, 3
'   '.InitEntrySounds "fx_drain", SoundFX(SSolenoidOn,DOFContactors), SoundFX(SSolenoidOn,DOFContactors)
'   '.InitExitSounds  SoundFX(SSolenoidOn,DOFContactors), SoundFX("fx_ballrel",DOFContactors)
'   .Balls = 4
'   .CreateEvents "bsTrough", Drain
'    End With


  set dtDrop=new cvpmDropTarget
    dtDrop.InitDrop Array(sw17,sw18,sw19,sw20,sw21,sw22,sw23),Array(17,18,19,20,21,22,23)
    dtDrop.InitSnd SoundFX("DTDrop",DOFContactors),SoundFX("DTReset",DOFContactors)

  set dtLL=new cvpmDropTarget
    dtLL.InitDrop sw52,52
    dtLL.InitSnd SoundFX("DTDrop",DOFContactors),SoundFX("DTReset",DOFContactors)

  CenterPost.IsDropped=1
  RightPost.IsDropped=1
  LeftPost.IsDropped=1

  Controller.Switch(9)=1 'playfield glass
  Controller.Switch(10)=0 'lockdown

End Sub

'**********************************************************************************************************
'Plunger code
'**********************************************************************************************************

Sub Table1_KeyDown(ByVal KeyCode)
  If KeyCode = LeftMagnaSave or KeyCode = KeyUpperLeft Then Controller.Switch(1)=1
  If KeyCode = RightMagnaSave or KeyCode = KeyUpperRight Then Controller.Switch(8)=1
' If KeyCode = KeyFront Then Controller.Switch()=1  'Tournament button
  If keycode = PlungerKey Then Plunger.Pullback:playsoundAtVol"plungerpull", Plunger, 1
  If KeyDownHandler(keycode) Then Exit Sub
End Sub

Sub Table1_KeyUp(ByVal KeyCode)
  If KeyCode = LeftMagnaSave or KeyCode = KeyUpperLeft Then Controller.Switch(1)=0
  If KeyCode = RightMagnaSave or KeyCode = KeyUpperRight Then Controller.Switch(8)=0

  If keycode = PlungerKey Then Plunger.Fire:PlaySoundAtVol"plunger", Plunger, 1
  If KeyUpHandler(keycode) Then Exit Sub
End Sub

Dim plungerIM

    ' Impulse Plunger
    Set plungerIM = New cvpmImpulseP
    With plungerIM
        .InitImpulseP swplunger, 55, 0.6
        .Random 0.3
        .switch 16
        .InitExitSnd SoundFX("fx_AutoPlunger",DOFContactors), SoundFX("fx_AutoPlunger",DOFContactors)
        .CreateEvents "plungerIM"
    End With

'**********************************************************************************************************

 ' Drain hole and kickers
Sub Drain_Hit:bsTrough.addball me : playsoundAtVol"drain" , Drain, 1: End Sub

'Drop Targets
Sub sw17_Dropped:dtDrop.Hit 1:End Sub
Sub sw18_Dropped:dtDrop.Hit 2:End Sub
Sub sw19_Dropped:dtDrop.Hit 3:End Sub
Sub sw20_Dropped:dtDrop.Hit 4:End Sub
Sub sw21_Dropped:dtDrop.Hit 5:End Sub
Sub sw22_Dropped:dtDrop.Hit 6:End Sub
Sub sw23_Dropped:dtDrop.Hit 7:End Sub

Sub sw52_Dropped:dtLL.Hit 1:End Sub

'Wire Triggers
Sub sw16_Hit:  playsoundAtVol"rollover" , ActiveBall, 1: End Sub 'Coded to Impulse Plunger
'Sub sw16_unHit:Controller.Switch(16)=0:End Sub
Sub sw31_Hit:Controller.Switch(31)=1 : playsoundAtVol"rollover" , ActiveBall, 1: End Sub
Sub sw31_unHit:Controller.Switch(31)=0:End Sub
Sub sw32_Hit:Controller.Switch(32)=1 : playsoundAtVol"rollover" , ActiveBall, 1: End Sub
Sub sw32_unHit:Controller.Switch(32)=0:End Sub
Sub sw33_Hit:Controller.Switch(33)=1 : playsoundAtVol"rollover" , ActiveBall, 1: End Sub
Sub sw33_unHit:Controller.Switch(33)=0:End Sub
Sub sw34_Hit:Controller.Switch(34)=1 : playsoundAtVol"rollover" , ActiveBall, 1: End Sub
Sub sw34_unHit:Controller.Switch(34)=0:End Sub
Sub sw35_Hit:Controller.Switch(35)=1 : playsoundAtVol"rollover" , ActiveBall, 1: End Sub
Sub sw35_unHit:Controller.Switch(35)=0:End Sub
Sub sw41_Hit:Controller.Switch(41)=1 : playsoundAtVol"rollover" , ActiveBall, 1: End Sub
Sub sw41_unHit:Controller.Switch(41)=0:End Sub
Sub sw57_Hit:Controller.Switch(57)=1 : playsoundAtVol"rollover" , ActiveBall, 1: End Sub
Sub sw57_unHit:Controller.Switch(57)=0:End Sub
Sub sw58_Hit:Controller.Switch(58)=1 : playsoundAtVol"rollover" , ActiveBall, 1: End Sub
Sub sw58_unHit:Controller.Switch(58)=0:End Sub
Sub sw60_Hit:Controller.Switch(60)=1 : playsoundAtVol"rollover" , ActiveBall, 1: End Sub
Sub sw60_unHit:Controller.Switch(60)=0:End Sub
Sub sw61_Hit:Controller.Switch(61)=1 : playsoundAtVol"rollover" , ActiveBall, 1: End Sub
Sub sw61_unHit:Controller.Switch(61)=0:End Sub

'Scoring Rubber
Sub sw24_hit:vpmTimer.pulseSw 24 : playsoundAtVol"flip_hit_3" , sw24, 1: End Sub

 'Stand Up Targets
Sub sw25_Hit:vpmTimer.PulseSw 25:sw25p.transY = -10:Me.TimerEnabled = 1:End Sub
Sub sw25_Timer:sw25p.transY = 0:Me.TimerEnabled = 0:End Sub
Sub sw26_Hit:vpmTimer.PulseSw 26:sw26p.transY = -10:Me.TimerEnabled = 1:End Sub
Sub sw26_Timer:sw26p.transY = 0:Me.TimerEnabled = 0:End Sub
Sub sw27_Hit:vpmTimer.PulseSw 27:sw27p.transY = -10:Me.TimerEnabled = 1:End Sub
Sub sw27_Timer:sw27p.transY = 0:Me.TimerEnabled = 0:End Sub
Sub sw28_Hit:vpmTimer.PulseSw 28:sw28p.transY = -10:Me.TimerEnabled = 1:End Sub
Sub sw28_Timer:sw28p.transY = 0:Me.TimerEnabled = 0:End Sub
Sub sw29_Hit:vpmTimer.PulseSw 29:sw29p.transY = -10:Me.TimerEnabled = 1:End Sub
Sub sw29_Timer:sw29p.transY = 0:Me.TimerEnabled = 0:End Sub
Sub sw30_Hit:vpmTimer.PulseSw 30:sw30p.transY = -10:Me.TimerEnabled = 1:End Sub
Sub sw30_Timer:sw30p.transY = 0:Me.TimerEnabled = 0:End Sub

Sub sw39_Hit:vpmTimer.PulseSw 39:sw39p.transY = -10:Me.TimerEnabled = 1:End Sub
Sub sw39_Timer:sw39p.transY = 0:Me.TimerEnabled = 0:End Sub
Sub sw40_Hit:vpmTimer.PulseSw 40:sw40p.transY = -10:Me.TimerEnabled = 1:End Sub
Sub sw40_Timer:sw40p.transY = 0:Me.TimerEnabled = 0:End Sub

'Ramp Triggers
Sub sw36_Hit:Controller.Switch(36)=1 : playsoundAtVol"rollover" , ActiveBall, 1: End Sub
Sub sw36_unHit:Controller.Switch(36)=0:End Sub
Sub sw42_Hit:Controller.Switch(42)=1 : playsoundAtVol"rollover" , ActiveBall, 1: End Sub
Sub sw42_unHit:Controller.Switch(42)=0:End Sub

'Bumpers
Sub Bumper1_Hit : vpmTimer.PulseSw(49) : playsoundAtVol SoundFX("fx_bumper1",DOFContactors), ActiveBall, 1: End Sub
Sub Bumper2_Hit : vpmTimer.PulseSw(51) : playsoundAtVol SoundFX("fx_bumper2",DOFContactors), ActiveBall, 1: End Sub
Sub Bumper3_Hit : vpmTimer.PulseSw(52) : playsoundAtVol SoundFX("fx_bumper3",DOFContactors), ActiveBall, 1: End Sub


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
     'If chgLamp(ii,0) = 2 Then solTrough chgLamp(ii,1)
     'If chgLamp(ii,0) = 4 Then PFGI chgLamp(ii,1)

        Next
    End If
    UpdateLamps
End Sub


Sub UpdateLamps()
NFadeL 1, Light1 'Ball 1
NFadeL 2, Light2 'Ball 2
NFadeL 3, Light3 'Ball 3
NFadeL 4, Light4 'Ball 4
NFadeL 5, Light5 'Ball 5
NFadeL 6, Light6 'Ball 6
NFadeL 7, Light7 'Ball 7
NFadeL 8, Light8 'Ball 8
NFadeL 9, Light9 'Ball 9
NFadeL 10, Light10 'Ball 10
NFadeL 11, Light11 'Ball 11
NFadeL 12, Light12 'Ball 12
NFadeL 13, Light13 'Ball 13
NFadeL 14, Light14 'Ball 14
NFadeL 15, Light15 'Ball 15

NFadeL 16, light16 'Shoot again

'Balls near the droptagets
NFadeL 17, Light17 'Ball 1
NFadeL 18, Light18 'Ball 2
NFadeL 19, light19 'Ball 3
NFadeL 20, Light20 'Ball 4
NFadeL 21, Light21 'Ball 5
NFadeL 22, Light22 'Ball 6
NFadeL 23, Light23 'Ball 7
NFadeL 24, Light24 'Ball 8
NFadeL 25, Light25 'Ball 9
NFadeL 26, Light26 'Ball 10
NFadeL 27, Light27 'Ball 11
NFadeL 28, Light28 'Ball 12
NFadeL 29, Light29 'Ball 13
NFadeL 30, Light30 'Ball 14
NFadeL 31, Light31 'Ball 15

NFadeL 32, Light32 '5x bottom playfield

NFadeL 33, Light33 '1M
NFadeL 34, Light34 '2M
NFadeL 35, Light35 '3M
NFadeL 36, Light36 '5M
NFadeL 37, Light37 'Extra Ball
NFadeL 38, Light38 'Special

NFadeL 39, Light39 'Special outlane left

NFadeL 40, Light40 '[V]ODKA

NFadeL 41, Light41 '2x
NFadeL 42, Light42 '3x
NFadeL 43, Light43 '4x
NFadeL 44, Light44 '5x
NFadeL 45, Light45 '5M
NFadeL 46, Light46 'Jackpot

NFadeL 47, Light47 'Goldbar 5M
NFadeL 48, Light48 'Goldbar Special

NFadeL 49, Light49 '[G]OLDEN
NFadeL 50, Light50 'G[O]LDEN
NFadeL 51, Light51 'GO[L]DEN
NFadeL 52, Light52 'GOL[D]EN
NFadeL 53, Light53 'GOLD[E]N
NFadeL 54, Light54 'GOLDE[N]

NFadeL 55, Light55 'VODK[A]

NFadeL 56, Light56 'Special outlane right
NFadeL 57, Light57 'Skill shot

'NFadeL 58, Light58 'Start button
'NFadeL 59, Light59 'Tournament Button

NFadeL 60, Light60 'Bank shot
NFadeL 61, Light61 'Trophy
NFadeL 62, Light62 'Lock
NFadeL 63, Light63 'Multi Ball
NFadeL 64, Light64 'Collect Bonus

NFadeL 65, Light65 'V[O]DKA
NFadeL 66, Light66 'VO[D]KA
NFadeL 67, Light67 'VOD[K]A
NFadeL 68, Light68 '2M near skillshot
NFadeL 69, Light69 '5M near skillshot
NFadeL 70, Light70 '2x bottom playfield
NFadeL 71, Light71 '3x bottom playfield
NFadeL 72, Light72 '4x bottom playfield
NFadeL 73, Light73 'Apron LED
NFadeL 74, Light74 'Apron LED
NFadeL 75, Light75 'Apron LED
NFadeL 76, Light76 'Apron LED
NFadeL 77, Light77 'Apron LED
NFadeL 78, Light78 'Apron LED
NFadeL 79, Light79 '8M triangle
NFadeL 80, Light80 '16M triangle



'Solenoid Controlled

'Flash Lower Right X3
NFadeLm 120, S120a
NFadeLm 120, S120b
NFadeLm 120, S120c
NFadeLm 120, S120d
NFadeLm 120, S120e
NFadeL 120, S120f

'Lower Left X4
NFadeLm 103, S103a
NFadeLm 103, S103b
NFadeLm 103, S103c
NFadeLm 103, S103d
NFadeLm 103, S103e
NFadeLm 103, S103f
NFadeLm 103, S103g
NFadeL 103, S103h

'Upper Left x2
NFadeLm 117, S117a
NFadeLm 117, S117b
NFadeLm 117, S117c
NFadeL 117, S117d

'Flash Uppper Right X3
NFadeLm 118, S118a
NFadeLm 118, S118b
NFadeLm 118, S118c
NFadeL 118, S118d

'Flash Mid Right X4
NFadeLm 119, S119a
NFadeLm 119, S119b
NFadeLm 119, S119c
NFadeLm 119, S119d
NFadeLm 119, S119e
NFadeLm 119, S119f
NFadeLm 119, S119g
NFadeL 119, S119h


'Flash Lower Right X3
'NFadeLm 120, S120a
'NFadeLm 120, S120b
'NFadeLm 120, S120c
'NFadeLm 120, S120d
'NFadeLm 120, S120e
'NFadeL 120, S120f

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
  vpmTimer.PulseSw 62
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
  vpmTimer.PulseSw 59
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
' Positional Sound Playback Functions by DJRobX, Rothbauerw, Thalamus and Herweh
' PlaySoundAtVol sound, 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 1, AudioFade(ActiveBall)
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


'********************************************************************
'      JP's VP10 Rolling Sounds (+rothbauerw's Dropping Sounds)
'********************************************************************

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

    For b = 0 to UBound(BOT)
        ' play the rolling sound for each ball
        If BallVel(BOT(b) ) > 1 AND BOT(b).z < 30 Then
            rolling(b) = True
            PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b)), AudioPan(BOT(b)), 0, Pitch(BOT(b)), 1, 0, AudioFade(BOT(b))
        Else
            If rolling(b) = True Then
                StopSound("fx_ballrolling" & b)
                rolling(b) = False
            End If
        End If

        ' play ball drop sounds
        If BOT(b).VelZ < -1 and BOT(b).z < 55 and BOT(b).z > 27 Then 'height adjust for ball drop sounds
            PlaySound "fx_ball_drop" & b, 0, ABS(BOT(b).velz)/17, AudioPan(BOT(b)), 0, Pitch(BOT(b)), 1, 0, AudioFade(BOT(b))
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
  FlipperLSh1.RotZ = LeftFlipper1.currentangle

'    LFLogo.RotY = LeftFlipper.CurrentAngle
'    RFlogo.RotY = RightFlipper.CurrentAngle
'    LFLogo1.RotY = LeftFlipper1.CurrentAngle


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

' Thalamus : Exit in a clean and proper way
Sub Table1_exit
  Controller.Pause = False
  Controller.Stop
End Sub

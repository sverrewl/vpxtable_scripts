'************
'************
'CREDITS
'************
'************

'Based On
' Stern's Cheetah / IPD No. 500 / June, 1980 / 4 Players
' VPX - version by JPSalas 2018, version 1.0.1
' Script based on destruk's script

' Thalamus 2018-07-20
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Changed UseSolenoids=1 to 2
' Thalamus 2018-08-11 : Improved directional sounds

'*********************************************************

'TIKI BOB'S EXOTIC BEACH PARTY (Original Theme)

'Created by iDigStuff
'2021, version 1.0.0
'2025, version 2.0.0

'Copiloted by Apophis
'B2S | Media | scripting help| fleep sounds | nFozzy physics | ???

'Special Thank You to VPin Workshop & Watacaractr & Scotty Wic
'Much appreciation JP Salas

'DMD Artwork Joey B
'Primitive assistance Andrei Maraklov
'Testing by Rik Laubach, Apophis

'FlexDMD added by endeemillr in honor of this amazing table

'***********************************************************
Option Explicit
Randomize

'***********************************************************
'Music Location
'***********************************************************

'https://mega.nz/folder/5Bp2TIpQ#RyB2iGHyGuOxHuxUh3iEgQ

'Unzip and put folder "tiki" into :\\visualpinball\music

'*************************************************************


'********************
'USER OPTIONS
'********************



Dim UseFlexDMD

UseFlexDMD = 1   ' 0 = OFF, 1 = ON


' Music Mod # of tracks
 'Place as many songs (MP3 ONLY) as you want in the Tiki music folder making sure to rename the files 1.mp3 2.mp3 3.mp3 etc...
 'Then come here and change this number to the total number of songs in the folder. All Done!

Const NumMusicTracks = 14



'********************
'END USER OPTIONS
'********************

'***************************************************************************


'///////////////////////-----General Sound Options-----///////////////////////
'// VolumeDial:
'// VolumeDial is the actual global volume multiplier for the mechanical sounds.
'// Values smaller than 1 will decrease mechanical sounds volume.
'// Recommended values should be no greater than 1.
Const VolumeDial = 1.5


'///////////////////////-----Phsyics Mods-----///////////////////////
Const RubberizerEnabled = 0     '0 = normal flip rubber, 1 = more lively rubber for flips
Const FlipperCoilRampupMode = 0     '0 = fast, 1 = medium, 2 = slow (tap passes should work)
Const TargetBouncerEnabled = 1    '0 = normal standup targets, 1 = bouncy targets
Const TargetBouncerFactor = 0.7   'Level of bounces. 0.2 - 1.5 are probably usable values.

Dim VRRoomChoice : VRRoomChoice = 2           '1 - Cab Only, 2 - Minimal, 3 - Tiki Room
Dim VRTest : VRTest = false
Dim VRRoom

Dim dspTriggered : dspTriggered = False
Sub Table1_OptionEvent(ByVal eventId)
    If eventId = 1 And Not dspTriggered Then dspTriggered = True : DisableStaticPreRendering = True : End If

  ' VR Room
    VRRoomChoice = Table1.Option("VR Room", 1, 3, 1, 3, 0, Array("VR Cab", "Minimal", "Tiki"))
  SetupRoom

    If eventId = 3 And dspTriggered Then dspTriggered = False : DisableStaticPreRendering = False : End If
End Sub

' VR Room Auto-Detect
Dim VR_Obj, VRMode

  If RenderingMode = 2 or VRTest Then
    VRMode = True
  Else
    VRMode = False
  End if

Sub SetupRoom
  for each VR_Obj in VRMinimalRoom:VR_Obj.visible = 0:Next
  for each VR_Obj in VR_Mega:VR_Obj.visible = 0:Next
  for each VR_Obj in VRCabinet:VR_Obj.visible = 0:Next
  vrBGAn1.enabled = 0
  If RenderingMode = 2 or VRTest Then
    VRRoom = VRRoomChoice
  Else
    VRRoom = 0
  End If
  If VRRoom > 0 Then
        flames1.rotx = -90: flames2.rotx = -90: flames3.rotx = -90: flames4.rotx = -90
        flames1.x = 870: flames2.x = 870: flames3.x = 870: flames4.x = 870
    lrail.Visible = 0
    rrail.Visible = 0
    lrail1.Visible = 0
    rrail1.Visible = 0
    ramp15.Visible = 0
    ramp34.Visible = 0
    Hippo001.visible = 0
    Hippo002.x = Hippo002.x-5.5
    Hippo.x = Hippo.x-5.5
    vrBGAn1.enabled = 1
    End If
  If VRRoom = 1 Then
    for each VR_Obj in VRMinimalRoom:VR_Obj.visible = 0:Next
    for each VR_Obj in VR_Mega:VR_Obj.visible = 0:Next
    for each VR_Obj in VRCabinet:VR_Obj.visible = 1:Next
    Pincab_Leg_Back_Left.z = 206.3
    Pincab_Leg_Back_Right.z = 206.3
  End If
  If VRRoom = 2 Then
    for each VR_Obj in VRMinimalRoom:VR_Obj.visible = 1:Next
    for each VR_Obj in VR_Mega:VR_Obj.visible = 0:Next
    for each VR_Obj in VRCabinet:VR_Obj.visible = 1:Next
    Pincab_Leg_Back_Left.z = 206.3
    Pincab_Leg_Back_Right.z = 206.3
  End If
  If VRRoom = 3 Then
    for each VR_Obj in VRMinimalRoom:VR_Obj.visible = 0:Next
    for each VR_Obj in VR_Mega:VR_Obj.visible = 1:Next
    for each VR_Obj in VRCabinet:VR_Obj.visible = 1:Next
    Pincab_Leg_Back_Left.z = 86.3
    Pincab_Leg_Back_Right.z = 86.3
  End If

End Sub

'********************
'Load Table
'********************

Const BallSize = 50
Const BallMass = 1

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

LoadVPM "01550000", "stern.vbs", 3.26

Dim bsTrough, bsSaucer, dtL, dtR, dtT, dtR2, x
Dim dtL_on   ' -apophis

Const cGameName = "cheetah"

Const UseSolenoids = 2
Const UseLamps = 0
Const UseGI = 0
Const UseSync = 0 'set it to 1 if the table runs too fast
Const HandleMech = 0



Dim VarHidden
If Table1.ShowDT = true then
    VarHidden = 1
    For each x in aReels
        x.Visible = 1
    Next
else
    VarHidden = 0
    For each x in aReels
        x.Visible = 0
    Next
    lrail.Visible = 0
    rrail.Visible = 0
    lrail1.Visible = 0
    rrail1.Visible = 0
end if

if B2SOn = true then VarHidden = 1


Dim GiLights


'************
' Table init.
'************

Sub table1_Init
  ' initalise the FlexDMD display
    If UseFlexDMD = 1 Then FlexDMD_Init
    vpmInit me
    With Controller
        .GameName = cGameName
        If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description:Exit Sub
        .SplashInfoLine = "Tiki Bob's Atomic Beach Party - by iDigStuff" & vbNewLine & "based on Cheetah Stern 1980 VPX table by JPSalas"
        .HandleKeyboard = 0
        .ShowTitle = 0
        .ShowDMDOnly = 1
        .ShowFrame = 0
        .HandleMechanics = 0
        .Hidden = VarHidden
        .Games(cGameName).Settings.Value("rol") = 0 '1= rotated display, 0= normal
    .Games(cGameName).Settings.Value("sound") = 0
    If UseFlexDMD = 1 Then ExternalEnabled = .Games(cGameName).Settings.Value("showpindmd")
    If UseFlexDMD = 1 Then .Games(cGameName).Settings.Value("showpindmd") = 0
        '.SetDisplayPosition 0,0, GetPlayerHWnd 'restore dmd window position
        On Error Resume Next
        Controller.SolMask(0) = 0
        vpmTimer.AddTimer 2000, "Controller.SolMask(0)=&Hffffffff'" 'ignore all solenoids - then add the Timer to renable all the solenoids after 2 seconds
        Controller.Run GetPlayerHWnd

        On Error Goto 0
    End With

    ' Nudging
    vpmNudge.TiltSwitch = 7
    vpmNudge.Sensitivity = 3
    vpmNudge.TiltObj = Array(Bumper1, Bumper2, Bumper3, LeftSlingshot, RightSlingshot, RightSlingshot1)

    ' Trough
    Set bsTrough = New cvpmBallStack
    With bsTrough
        .InitNoTrough BallRelease, 33, 90, 4
        '.InitExitSnd SoundFX("fx_ballrel", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)  'apophis
        .Balls = 1
    End With

    ' Saucers
    Set bsSaucer = New cvpmBallStack
    bsSaucer.InitSaucer sw32, 32, 0, 25
    'bsSaucer.InitExitSnd SoundFX("fx_kicker", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)  'apophis
    bsSaucer.KickForceVar = 6

    ' Drop targets
  dtL_on = 0   ' -apophis

    Set dtL = New cvpmDropTarget
    dtL.InitDrop Array(sw24, sw23, sw22, sw21, sw20), Array(24, 23, 22, 21, 20)
    dtL.initsnd SoundFX("", DOFDropTargets), SoundFX("fx_resetdrop", DOFContactors)
    dtL.CreateEvents "dtL"

    Set dtT = New cvpmDropTarget
    dtT.InitDrop Array(sw25, sw26, sw27), Array(25, 26, 27)
    dtT.initsnd SoundFX("", DOFDropTargets), SoundFX("fx_resetdrop", DOFContactors)
    dtT.CreateEvents "dtT"

    Set dtR2 = New cvpmDropTarget
    dtR2.InitDrop Array(sw28, sw29, sw30), Array(28, 29, 30)
    dtR2.initsnd SoundFX("", DOFDropTargets), SoundFX("fx_resetdrop", DOFContactors)
    dtR2.CreateEvents "dtR2"

    Set dtR = New cvpmDropTarget
    dtR.InitDrop Array(sw17, sw18, sw19), Array(17, 18, 19)
    dtR.initsnd SoundFX("", DOFDropTargets), SoundFX("fx_resetdrop", DOFContactors)
    dtR.CreateEvents "dtR"

    ' Main Timer init
    PinMAMETimer.Interval = PinMAMEInterval
    PinMAMETimer.Enabled = 1

    ' Turn on Gi
    GiOff

  flames1.opacity = 0
  flames2.opacity = 0
  flames3.opacity = 0
  flames4.opacity = 0


  tbeam0.opacity = 0
  tbeam1.opacity = 0
  tbeam2.opacity = 0
  tbeam3.opacity = 0
  tbeam4.opacity = 0
  tbeam5.opacity = 0
  tbeam6.opacity = 0

  If VRMode = True Then
    setup_backglass
    If UseFlexDMD = 1 Then
      DMD.visible = 1
      VRFlexDMDTopper.visible = 1
    Else
      DMD.visible = 0
      VRFlexDMDTopper.visible = 0
    End if
  End If

End Sub

Sub table1_Paused: Controller.Pause = 1: End Sub
Sub table1_unPaused: Controller.Pause = 0: End Sub
Sub table1_Exit
  Controller.Games(cGameName).Settings.Value("sound") = 1
  If UseFlexDMD = 1 then
    If Not FlexDMD is Nothing Then
      FlexDMD.Show = False
      FlexDMD.Run = False
      FlexDMD = NULL
      Controller.Games(cGameName).Settings.Value("showpindmd")=ExternalEnabled
    End if
  End if
  Controller.Stop
End Sub


Sub MusicOn
    Dim x
    x = INT(NumMusicTracks * Rnd + 1)
      PlayMusic "Tiki\"& x &".mp3"
 End Sub

 Sub Table1_MusicDone()
    MusicOn
 End Sub

Dim Track
Sub NextSong

If Track = 0 Then PlayMusic"Tiki\1.mp3" End If
If Track = 1 Then PlayMusic"Tiki\2.mp3" End If
If Track = 2 Then PlayMusic"Tiki\3.mp3" End If
If Track = 3 Then PlayMusic"Tiki\4.mp3" End If
If Track = 4 Then PlayMusic"Tiki\5.mp3" End If
If Track = 5 Then PlayMusic"Tiki\6.mp3" End If
If Track = 6 Then PlayMusic"Tiki\7.mp3" End If
If Track = 7 Then PlayMusic"Tiki\8.mp3" End If
If Track = 8 Then PlayMusic"Tiki\9.mp3" End If
If Track = 9 Then PlayMusic"Tiki\10.mp3" End If
If Track = 10 Then PlayMusic"Tiki\11.mp3" End If
If Track = 11 Then PlayMusic"Tiki\12.mp3" End If
If Track = 12 Then PlayMusic"Tiki\13.mp3" End If
If Track = 13 Then PlayMusic"Tiki\14.mp3" End If


Track = (Track + 1) mod 14
End Sub



'**********
' Keys
'**********

Sub table1_KeyDown(ByVal Keycode)

  If Keycode = LeftFlipperKey Then
    Pincab_Flipper_Left.X = Pincab_Flipper_Left.X + 10
  End If

  If Keycode = RightFlipperKey Then
        Pincab_Flipper_Right.X = Pincab_Flipper_Right.X - 10
  End If

  If keycode = PlungerKey Then
    TimerPlunger.Enabled = True
    TimerPlunger2.Enabled = False
  End If

  if keycode=StartGameKey then
    Pincab_Start_Button.y = Pincab_Start_Button.y -3
  End If

  If KeyCode = RightMagnaSave AND GiLights = 1 Then NextSong
    If KeyCode = LeftMagnaSave Then EndMusic
  If keycode = StartGameKey Then soundStartButton():MusicOn: credittimer.Enabled = false: creditframe = -400: digitson
  If keycode = StartGameKey AND GiLights=0 Then EndMusic
    If keycode = LeftTiltKey Then Nudge 90, 2:SoundNudgeLeft()
    If keycode = RightTiltKey Then Nudge 270, 2:SoundNudgeRight()
    If keycode = CenterTiltKey Then Nudge 0, 3:SoundNudgeCenter()
    If keycode = PlungerKey And GiLights = 0 Then SoundPlungerPull():Plunger.Pullback
    If keycode = PlungerKey And GiLights = 1 Then SoundPlungerPull():Plunger.Pullback : PlaySound"alien5"
  If keycode = keyInsertCoin1 or keycode = keyInsertCoin2 or keycode = keyInsertCoin3 or keycode = keyInsertCoin4 Then
    Select Case Int(rnd*3)
      Case 0: PlaySound ("Coin_In_1"), 0, CoinSoundLevel, 0, 0.25
      Case 1: PlaySound ("Coin_In_2"), 0, CoinSoundLevel, 0, 0.25
      Case 2: PlaySound ("Coin_In_3"), 0, CoinSoundLevel, 0, 0.25
    End Select
  StopSound"1"
  StopSound"2"
  StopSound"3"
  StopSound"4"
  StopSound"5"
  Coinin

  End If
      If vpmKeyDown(keycode) Then Exit Sub
End Sub

Sub table1_KeyUp(ByVal Keycode)

  If Keycode = LeftFlipperKey Then
    Pincab_Flipper_Left.X = Pincab_Flipper_Left.X - 10
  End If

  If Keycode = RightFlipperKey Then
    Pincab_Flipper_Right.X = Pincab_Flipper_Right.X + 10
  End If

  If KeyCode = PlungerKey Then
    TimerPlunger.Enabled = False
    TimerPlunger2.Enabled = True
    Pincab_Plunger.Y = 2241.833
  End If

  if keycode=StartGameKey then
    Pincab_Start_Button.y = Pincab_Start_Button.y +3
  End If

    If keycode = PlungerKey And GiLights=0 Then SoundPlungerReleaseBall():Plunger.Fire
    If keycode = PlungerKey And GiLights=1 Then SoundPlungerReleaseBall():Plunger.Fire:StopSound"alien5":PlaySound"ignition": moviemode = 13: digitsoff: framenumber=0
    If vpmKeyUp(keycode) Then Exit Sub
End Sub

Sub Delay( seconds )
  Dim wshShell, strCmd
  Set wshShell = CreateObject( "WScript.Shell" )
  strCmd = wshShell.ExpandEnvironmentStrings( "%COMSPEC% /C (PING.EXE -n " & ( seconds + 1 ) & " localhost >NUL 2>&1)" )
  wshShell.Run strCmd, 0, 1
  Set wshShell = Nothing
End Sub

Sub Gate2_hit
musicOn
DrainNotRed
GiMotelTiki.BlinkPattern = 00010
GiMotelTiki.BlinkInterval = 1000
GiMotelTiki.state=2
end sub


'************
'Sound Subs
'************
sub coinin
Dim yy
  yy = INT(5 * RND(1) )
  Select Case yy
    Case 0:PlaySound"1"
    Case 1:PlaySound"2"
    Case 2:PlaySound"3"
    Case 3:PlaySound"4"
    Case 4:PlaySound"5"
    End Select
End Sub


Sub Bumperhit
  Dim z
  z = INT(6 * RND(1) )
  Select Case z
    Case 0:PlaySound"birdcalls1",0,.3
    Case 1:PlaySound"birdcalls2",0,.3
    Case 2:PlaySound"birdcalls3",0,.3
    Case 3:PlaySound"birdcalls4",0,.3
    Case 4:PlaySound"birdcalls5",0,.3
    Case 5:PlaySound"birdcalls6",0,.3
  End Select
End Sub

Sub cowbeam
  Dim bb
  bb = INT(4 * RND(1) )
  Select Case bb
    Case 0:PlaySound"cow0"
    Case 1:PlaySound"cow1"
    Case 2:PlaySound"cow2"
    Case 3:PlaySound"cow3"
  End Select
End Sub

Sub k_hit
  Dim zz
  zz = INT(17 * RND(1) )
  Select Case zz
    Case 0:PlaySound"knewsflash"
    Case 1:PlaySound"kcountdown"
    Case 2:PlaySound"kaliencow"
    Case 3:PlaySound"kInvaders"
    Case 4:PlaySound"kbeatmachine"
    Case 5:PlaySound"spaceship"
    Case 6:PlaySound"scifi"
    Case 7:PlaySound"gunstanksbombs"
    Case 8:PlaySound"beginningoftheend"
    Case 9:PlaySound"duetocircumstances"
    Case 10:PlaySound"fieldtransmission"
    Case 11:PlaySound"extinct"
    Case 12:PlaySound"stayhome"
    Case 13:PlaySound"conditioned"
    Case 14:PlaySound"nineminutes"
    Case 15:PlaySound"flyingsaucers"
    Case 16:PlaySound"phenomenom"

  End Select
End Sub

Sub aliens
  Dim xx
  xx =  INT(4 * RND(1) )
  select case xx
    Case 0: PlaySound"alien3"
    Case 1: PlaySound"alien4"
    Case 2: PlaySound"alien5"
    Case 3: PlaySound"whip"

  End Select
End Sub


sub launchsound
  Dim y
  y= INT (4 * RND(1) )
  select case y
    Case 0: PlaySound"beam"
    Case 1: PlaySound"alien3"
    Case 2: PlaySound"rocket"
    Case 3: PlaySound"bionic"
  End Select
End Sub


'*************
' Slingshots
'*************

Dim LStep, RStep, RStep1

Sub LeftSlingShot_Slingshot
    RandomSoundSlingshotLeft Lemk
  PlaySound"drum hits2"
  'aliens
    LeftSling4.Visible = 1
    Lemk.RotX = 26
    LStep = 0
    vpmTimer.PulseSw 16
    LeftSlingShot.TimerEnabled = 1
  Gi2.duration 0,150,1
  Gi3.duration 0,150,1

End Sub


Sub LeftSlingShot_Timer
    Select Case LStep
        Case 1:LeftSling4.Visible = 0:LeftSLing3.Visible = 1:Lemk.RotX = 14
        Case 2:LeftSLing3.Visible = 0:LeftSLing2.Visible = 1:Lemk.RotX = 2
        Case 3:LeftSLing2.Visible = 0:Lemk.RotX = -20:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

Sub RightSlingShot_Slingshot
    RandomSoundSlingshotRight Remk
  PlaySound"drum hits2"
' aliens
    RightSling4.Visible = 1
    Remk.RotX = 26
    RStep = 0
    vpmTimer.PulseSw 11
    RightSlingShot.TimerEnabled = 1
  Gi4.duration 0,150,1
  Gi5.duration 0,150,1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 1:RightSLing4.Visible = 0:RightSLing3.Visible = 1:Remk.RotX = 14
        Case 2:RightSLing3.Visible = 0:RightSLing2.Visible = 1:Remk.RotX = 2
        Case 3:RightSLing2.Visible = 0:Remk.RotX = -20:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

Sub RightSlingShot1_Slingshot
    RandomSoundSlingshotRight Remk1
  aliens
    RightSling8.Visible = 1
    Remk1.RotX = 26
    RStep1 = 0
    vpmTimer.PulseSw 15
    RightSlingShot1.TimerEnabled = 1
  Gi6.duration 2,1500,1
  Gi7.duration 2,1500,1
  digitsoff
  framenumber=0
  hidebirds
  moviemode = 19
End Sub

Sub RightSlingShot1_Timer
    Select Case RStep1
        Case 1:RightSLing8.Visible = 0:RightSLing7.Visible = 1:Remk1.RotX = 14
        Case 2:RightSLing7.Visible = 0:RightSLing6.Visible = 1:Remk1.RotX = 2
        Case 3:RightSLing6.Visible = 0:Remk1.RotX = -20:RightSlingShot1.TimerEnabled = 0
    End Select
    RStep1 = RStep1 + 1
End Sub

' Rubber & animations
Dim Rub1

'******************
'Tiki Motel
'******************

Sub sw40_slingshot
  RandomSoundSlingshotLeft Lemk1
  PlaySound"bulbout"
  vpmTimer.PulseSw 40
  Rub1 = 1
  sw40_Timer
  gimoteltiki.state=2
  gimoteltiki.BlinkPattern = 0000000010001000010101000110000000000000000010010010001
  gimoteltiki.BlinkInterval = 25
  moteltimer.enabled = 1
  moviemode=14
  digitsoff
  framenumber=0
End Sub


Sub sw40_Timer
    Select Case Rub1
        Case 1:r1.Visible = 1:sw40.TimerEnabled = 1
        Case 2:r1.Visible = 0:r2.Visible = 1
        Case 3:r2.Visible = 0:sw40.TimerEnabled = 0
    End Select
    Rub1 = Rub1 + 1
End Sub

'*******************
'Bumpers Animation
'*******************

Dim dirRing1:dirRing1 = -1
Dim dirRing2:dirRing2 = -1
Dim dirRing3:dirRing3 = -1

' Bumpers
Sub Bumper1_Hit:vpmTimer.PulseSw 12:RandomSoundBumperTop Bumper1:Gi35.duration 0,150,1:Gi34.duration 0,150,1:BumperHit:Bumper1.TimerEnabled=1
 If moviemode < 6 or moviemode = spriterows or moviemode = birdmode Then
  moviemode = birdmode
  bump1frame=0
  digitsoff
  framenumber=0
  movemovie
  End If
End Sub
Sub Bumper2_Hit:vpmTimer.PulseSw 13:RandomSoundBumperMiddle Bumper2:Gi33.duration 0,150,1:Gi32.duration 0,150,1:Bumperhit:Bumper2.TimerEnabled=1
 If moviemode < 6 or moviemode = spriterows or moviemode = birdmode  Then
  bump2frame=0
  moviemode = birdmode
  digitsoff
  framenumber=0
  movemovie
  End If
End Sub
Sub Bumper3_Hit:vpmTimer.PulseSw 14:RandomSoundBumperBottom Bumper3:Gi22.duration 0,150,1:Gi18.duration 0,150,1:Bumperhit:Bumper3.TimerEnabled=1
  If moviemode < 6 or moviemode = spriterows  or moviemode = birdmode Then
  bump3frame=0
  moviemode = birdmode
  digitsoff
  framenumber=0
  movemovie
  End If
End Sub


Sub Bumper1_timer()
  BumperRing1.Z = BumperRing1.Z + (5 * dirRing1)
  If BumperRing1.Z <= -40 Then dirRing1 = 1
  If BumperRing1.Z >= 0 Then
    dirRing1 = -1
    BumperRing1.Z = 0
    Me.TimerEnabled = 0
  End If
End Sub

Sub Bumper2_timer()
  BumperRing2.Z = BumperRing2.Z + (5 * dirRing2)
  If BumperRing2.Z <= -40 Then dirRing2 = 1
  If BumperRing2.Z >= 0 Then
    dirRing2 = -1
    BumperRing2.Z = 0
    Me.TimerEnabled = 0
  End If
End Sub

Sub Bumper3_timer()
  BumperRing3.Z = BumperRing3.Z + (5 * dirRing3)
  If BumperRing3.Z <= -40 Then dirRing3 = 1
  If BumperRing3.Z >= 0 Then
    dirRing3 = -1
    BumperRing3.Z = 0
    Me.TimerEnabled = 0
  End If
End Sub


'*********************
'Drain & Alien Attack
'*********************

Sub Drain_Hit
  RandomSoundDrain Drain
  bsTrough.AddBall Me
  EndMusic
  GiRedOff
  hidebirds
  framenumber=0
  digitsoff
  If girl = true then
  moviemode = 16:giggle
  Else if cow = true then
  moviemode = 17 :moo
  Else moviemode = 15
  End If
  End If
End Sub



'Sub sw32_Hit
' moviemode= 6 : girl = True:hidebirds: digitsoff: framenumber=0 :SoundSaucerLock:vpmTimer.AddTimer 1000,"bsSaucer.AddBall 0'": EndMusic : GiAlien : k_hit : End Sub



 Sub sw32_Hit
     moviemode = 6
     girl = True
     hidebirds
     digitsoff
     framenumber = 0    ' (keep your existing framenumber logic here – just remove the bsSaucer.AddBall string)
     EndMusic
     StartAlienMode        ' <<< new: call our helper sub
     k_hit
 End Sub

   '————————————————————————————————————————————
   ' Begin deferred Alien‑mode helper
   Sub StartAlienMode
    GiAlien
    AlienTimer.Interval = 1200   ' delay in milliseconds
    AlienTimer.Enabled  = True
   End Sub
   '————————————————————————————————————————————

Sub AlienTimer_Timer()
    AlienTimer.Enabled = False   ' stop the timer from repeating
    bsSaucer.AddBall 0           ' now eject the ball
End Sub

'*****************
'Drop Targets
'*****************

'Sub DTR test Sounds

Sub sw17_Hit:Controller.Switch(17) = 1 :Playsound"rumble":volcanotrigger: End Sub
Sub sw17_UnHit:Controller.Switch(17) = 0:End Sub

Sub sw18_Hit:Controller.Switch(18) = 1 : Playsound"rumble":volcanotrigger:End Sub
Sub sw18_UnHit:Controller.Switch(18) = 0:End Sub

Sub sw19_Hit:Controller.Switch(19) = 1 : Playsound"rumble":volcanotrigger:End Sub
Sub sw19_UnHit:Controller.Switch(19) = 0:End Sub


'Sub DTL test Sounds

Sub sw20_Hit
  Controller.Switch(20) = 1
  If dtL_on=5 or li55.state=1 then    '-apophis
        dtL_on = 1
    PlaySound"successa"
    PlaySound"tikihit1"
    GiFlash
  Else
    PlaySound"tikihit2"
  End If
End Sub
Sub sw20_UnHit:Controller.Switch(20) = 0:End Sub


Sub sw21_Hit
  Controller.Switch(21) = 1
  If dtL_on=4 or li8.state=1  then    '-apophis
    PlaySound"successb"
    PlaySound"tikihit1"
    GiFlash
  Else
    PlaySound"tikihit2"
  End If
End Sub
Sub sw21_UnHit:Controller.Switch(21) = 0:End Sub


Sub sw22_Hit
Controller.Switch(22) = 1
  If dtL_on=3 or li24.state=1  then    '-apophis
    PlaySound"successb"
    PlaySound"tikihit1"
    GiFlash
  Else
    PlaySound"tikihit2"
  End If
End Sub
Sub sw22_UnHit:Controller.Switch(22) = 0:End Sub

Sub sw23_Hit
  Controller.Switch(23) = 1
  If dtL_on=2 or li40.state=1  then    '-apophis
    PlaySound"successb"
    PlaySound"tikihit1"
    GiFlash
  Else
    PlaySound"tikihit2"
  End If
End Sub
Sub sw23_UnHit:Controller.Switch(23) = 0:End Sub

Sub sw24_Hit
  Controller.Switch(24) = 1
  If dtL_on=1 or li56.state=1  then    '-apophis
    PlaySound"successb"
    PlaySound"tikihit1"
    GiFlash
  Else
    PlaySound"tikihit2"
  End If
End Sub
Sub sw24_UnHit:Controller.Switch(24) = 0:End Sub


'Sub DTT test Sounds

Sub sw25_Hit:Controller.Switch(25) = 1 : PlaySound"bongos":End Sub
Sub sw25_UnHit:Controller.Switch(25) = 0:End Sub

Sub sw26_Hit:Controller.Switch(26) = 1 : PlaySound"bongos": End Sub
Sub sw26_UnHit:Controller.Switch(26) = 0:End Sub

Sub sw27_Hit:Controller.Switch(27) = 1 : PlaySound"bongos":End Sub
Sub sw27_UnHit:Controller.Switch(27) = 0:End Sub


'Sub DTR test Sounds

Sub sw28_Hit:Controller.Switch(28) = 1 : PlaySound"bongos": End Sub
Sub sw28_UnHit:Controller.Switch(28) = 0:End Sub

Sub sw29_Hit:Controller.Switch(29) = 1 : PlaySound"bongos": End Sub
Sub sw29_UnHit:Controller.Switch(29) = 0:End Sub

Sub sw30_Hit:Controller.Switch(30) = 1 : PlaySound"bongos": End Sub
Sub sw30_UnHit:Controller.Switch(30) = 0:End Sub


'**************
' Rollovers
'**************
Sub sw35_Hit:Controller.Switch(35) = 1: PlaySound "gong": DrainRed :KillGi:End Sub
Sub sw35_UnHit:Controller.Switch(35) = 0:End Sub

Sub sw34_Hit:Controller.Switch(34) = 1: PlaySound "gong":DrainRed: KillGi: End Sub
Sub sw34_UnHit:Controller.Switch(34) = 0:End Sub

'idigstuff
' ===== Shared handlers =====
Sub HandleInlaneHit(sw)
  Controller.Switch(sw) = 1
  If ActiveBall.VelY > 0 Then      ' toward flippers (swap if reversed on your table)
    PlaySound "hit"
  Else                              ' away from flippers
    PlaySound "unhit"
  End If
End Sub

Sub HandleInlaneUnHit(sw)
  Controller.Switch(sw) = 0
End Sub

' ===== Wrappers per trigger (minimal boilerplate) =====
Sub sw37_Hit():   HandleInlaneHit 37:   End Sub
Sub sw37_UnHit(): HandleInlaneUnHit 37: End Sub

Sub sw39_Hit():   HandleInlaneHit 39:   End Sub
Sub sw39_UnHit(): HandleInlaneUnHit 39: End Sub

Sub sw38_Hit():   HandleInlaneHit 38:   End Sub
Sub sw38_UnHit(): HandleInlaneUnHit 38: End Sub

Sub sw36_Hit():   HandleInlaneHit 36:   End Sub
Sub sw36_UnHit(): HandleInlaneUnHit 36: End Sub



'*********************
'   Radio treadmill
'**********************

Sub sw31_Hit:Controller.Switch(31) = 1
'PlaySound "whoosh1"
'PlaySound"whip"
'EndMusic
'sfxon
'musicon
end sub


sub sfxon
  Dim t
  t= INT (5 * RND(1) )
  select case t
  Case 0: PlaySound"1"
  Case 1: PlaySound"2"
  Case 2: PlaySound"3"
  Case 3: PlaySound"4"
  Case 4: PlaySound"5"

End Select
End Sub



Sub sw31_UnHit:Controller.Switch(31) = 0:End Sub


'***********************************************


' =========================
' Spinners (direction-aware for sw5)
' =========================

' Optional: keep these near your other globals if you want to tweak later
Dim sw5LastDir : sw5LastDir = 0   ' -1 or +1 (not strictly required, but handy if you want debouncing)
Const SW5_SPEED_MIN = 5           ' ignore tiny nudges that cause jitter

Sub sw4_Spin
    vpmTimer.PulseSw 4
    PlaySoundAtVol "fx_spinner", sw4, SpinSoundLevel
    PlaySound "sonarpulse"
End Sub

Sub sw5_Spin
    vpmTimer.PulseSw 5
    PlaySoundAtVol "fx_spinner", sw5, SpinSoundLevel

    Dim vx, vy, dir, speed
    vx = ActiveBall.VelX
    vy = ActiveBall.VelY

    ' Pick the axis with the stronger push
    If Abs(vx) >= Abs(vy) Then
        dir = Sgn(vx)        ' +1 or -1
        speed = Abs(vx)
    Else
        dir = Sgn(vy)
        speed = Abs(vy)
    End If

    ' Ignore near-zero wiggles
    If speed < SW5_SPEED_MIN Then Exit Sub

    ' Map directions:
    '  dir >= 0  -> "sonarpulse"
    '  dir <  0  -> "hippos" then "sonarpulse"
    If dir >= 0 Then
        PlaySound ""
    Else
        PlaySound "hippos"
        PlaySound ""
    End If

    sw5LastDir = dir
End Sub

Sub sw9_Spin
    vpmTimer.PulseSw 9
    PlaySoundAtVol "fx_spinner", sw9, SpinSoundLevel
End Sub

' Removed per request:
' Sub hippotr_hit: PlaySound "hippos": End Sub




'***********************
'Volcano
'***********************
Sub volcanotr_hit:volcanotrigger:fireball.enabled =  true : PlaySound"volcanochant":Playsound"explosion":Playsound"Rumble":moviemode= 11 :hidebirds: digitsoff: framenumber=0:End Sub

sub volcanotrigger
  v1.state=2
  v2.state=2
  v3.state=2
  v4.state=2
  v1.BlinkPattern = 000000001000100001010100
  v1.BlinkInterval = 40
  v2.BlinkPattern = 000000001000100001010100
  v2.BlinkInterval = 50
  v3.BlinkPattern = 000000001000100001010100
  v3.BlinkInterval = 50
  v4.BlinkPattern = 000000001000100001010100
  v4.BlinkInterval = 40
  volcanotimer.interval = 2800
  volcanotimer.enabled=1



  giBeachL.state=0
  giBeachR.state=0
  giBeachN.state=0
  giBeachS.state=0

  KillGi
  GiRed

  giOceanLRed.state=2
  giOceanLRed.BlinkPattern = 1101101100
  giOceanLRed.BlinkInterval = 110
  giOceanNRed.state=2
  giOceanNRed.BlinkPattern = 1101101100
  giOceanNRed.BlinkInterval =90
  giOceanRRed.state=2
  giOceanRRed.BlinkPattern = 1101101100
  giOceanRRed.BlinkInterval = 100
  giOceanSRed.state=2
  giOceanSRed.BlinkPattern = 1101101100
  giOceanSRed.BlinkInterval = 85

End Sub

Sub VolcanoTimer_Timer()
  VolcanoTimer.enabled=0
  v1.state=0
  v2.state=0
  v3.state=0
  v4.state=0
  giOceanSRed.state=0
  giOceanNRed.state=0
  giOceanRRed.state=0
  giOceanLRed.state=0

  if gilights=1 then

  giBeachL.state=1
  giBeachR.state=1
  giBeachN.state=1
  GiRedOff
  GiOn

End If
End Sub


' Volcano Fire
  dim firepos:firepos = 0
  sub fireball_timer
    flames1.opacity = 0
    flames2.opacity = 0
    flames3.opacity = 0
    flames4.opacity = 0
    firepos = firepos + 1
    select case firepos
      case 1 : flames1.opacity = 1000
      case 2 : flames2.opacity = 1000
      case 3 : flames3.opacity = 1000
      case 4 : flames4.opacity = 1000
      case 5 : flames2.opacity = 1000
      case 6 : flames3.opacity = 1000
      case 7 : flames4.opacity = 1000
      case 8 : flames2.opacity = 1000
      case 9 : flames3.opacity = 1000
      case 10 : flames4.opacity = 1000
      case 11 : flames2.opacity = 1000
      case 12 : flames3.opacity = 1000
      case 13 : flames4.opacity = 1000
      case 14 : flames4.opacity = 0 : firepos = 0 : fireball.enabled = false
    end Select
  end Sub

'****************************************************
'****************
'Lightning
'****************

'Targets
Sub sw10_Hit:vpmTimer.PulseSw 10
EndMusic
moviemode=12 :digitsoff: hidebirds: framenumber=0
StopSound "Thunder"
StopSound "chant"
PlaySound "Thunder"
PlaySound "chant"
'GiFlash2
GiLightning
LightningTimer.interval= 3000
LightningTimer.enabled =1
End Sub


Sub LightningTimer_Timer()
LightningTimer.enabled=0
For each x in aGiLights
        x.State = 1
next
For each x in aFxNight
        x.State = 0

  fx1.state=0
  fx2.state=0
  fx3.state=0
next
MusicOn
End Sub

sub TriggerTbeam_Hit
  tbeam.enabled = 1
end sub

sub Orbit_hit
  tractorpull=True
  PlaySound"beammeup"
  cowbeam
  UFOLights
  moviemode= 9
  hidebirds
  cow=true
  framenumber=0
  digitsoff


End Sub




'**************
'Tractor Beam
'**************

sub TriggerTractorBeam_hit
  LightSeqBeam.UpdateInterval = 1
  LightSeqBeam.Play SeqDownOn, 0, 0, 3900
  LightSeqBeam.Play SeqUpOff, 0, 0
  Playsound"ktractorbeam"

End Sub

Sub LightSeqBeam_PlayDone()
  LightSeqBeam.Play SeqAllOff

End Sub


sub starlight
  For each x in stars
    x.state = 2
next
End Sub


sub nostars
 For each x in stars
  x.state = 0
next
end sub



' UFO Tractor Beam Rings (apophis)
dim tractorpull:tractorpull = False
dim tbeamcnt:tbeamcnt = 0
const tbeamcntmax = 300
const tbeamopmax = 600
const tringystart = 25



sub tbeam_timer
  dim BOT
  tbeamcnt = tbeamcnt + 1
  'fade main beam
  if tbeamcnt < tbeamcntmax-200 Then
    tbeam0.opacity = tbeamopmax*min(1,tbeam0.opacity/tbeamopmax + 0.01)
  else
    tbeam0.opacity = tbeamopmax*max(0,tbeam0.opacity/tbeamopmax - 0.03)
  end if
  'fade rings in and out based on ball y position
  if tractorpull then
    BOT = GetBalls
    if BOT(0).y < tbeam1.y+tringystart then tbeam1.opacity = tbeamopmax*DistanceFade(tbeam1.y+tringystart-BOT(0).y,200)
    if BOT(0).y < tbeam2.y+tringystart then tbeam2.opacity = tbeamopmax*DistanceFade(tbeam2.y+tringystart-BOT(0).y,190)
    if BOT(0).y < tbeam3.y+tringystart then tbeam3.opacity = tbeamopmax*DistanceFade(tbeam3.y+tringystart-BOT(0).y,180)
    if BOT(0).y < tbeam4.y+tringystart then tbeam4.opacity = tbeamopmax*DistanceFade(tbeam4.y+tringystart-BOT(0).y,170)
    if BOT(0).y < tbeam5.y+tringystart then tbeam5.opacity = tbeamopmax*DistanceFade(tbeam5.y+tringystart-BOT(0).y,160)
    if BOT(0).y < tbeam6.y+tringystart then tbeam6.opacity = tbeamopmax*DistanceFade(tbeam6.y+tringystart-BOT(0).y,150)
  end if

  'disable timer when tractor beam event is done
  if tbeamcnt >= tbeamcntmax then
    tbeamcnt = 0
    tbeam.enabled=false
    tractorpull = false
    tbeam0.opacity = 0
    tbeam1.opacity = 0
    tbeam2.opacity = 0
    tbeam3.opacity = 0
    tbeam4.opacity = 0
    tbeam5.opacity = 0
    tbeam6.opacity = 0
  end if
end Sub

sub gate1_hit: tractorpull=false: end sub

function DistanceFade(dist,dpeak)
  DistanceFade  = max(0,dist/dpeak - 2*max(0,(dist-dpeak)/dpeak))
end function

Function max(a,b)
  if a > b then
    max = a
  Else
    max = b
  end if
end Function

Function min(a,b)
  if a > b then
    min = b
  Else
    min = a
  end if
end Function




'**************
'UFO Light Sequence
'**************

sub UFOLights

  UFOSeq.UpdateInterval = 3
  UFOSeq.Play SeqScrewLeftOn, 90, 5

End Sub

Sub UFOSeq_PlayDone()
  GiOn

End Sub


'**************
'speedbelt
'**************
dim dmdcycle
sub KickBoost_hit
KickBoost.Kick 0, 25
end sub

' speed belt
  dim beltpos
  Sub Beltmove_Timer
    Select Case beltpos
      Case 1:belt1.opacity = 50:belt2.opacity = 0:belt3.opacity = 0:belt4.opacity = 0:belt5.opacity = 0:belt6.opacity = 0
      Case 2:belt1.opacity = 0:belt2.opacity = 50:belt3.opacity = 0:belt4.opacity = 0:belt5.opacity = 0:belt6.opacity = 0
      Case 3:belt1.opacity = 0:belt2.opacity = 0:belt3.opacity = 50:belt4.opacity = 0:belt5.opacity = 0:belt6.opacity = 0
      Case 4:belt1.opacity = 0:belt2.opacity = 0:belt3.opacity = 0:belt4.opacity = 50:belt5.opacity = 0:belt6.opacity = 0
      Case 5:belt1.opacity = 0:belt2.opacity = 0:belt3.opacity = 0:belt4.opacity = 0:belt5.opacity = 50:belt6.opacity = 0
      Case 6:belt1.opacity = 0:belt2.opacity = 0:belt3.opacity = 0:belt4.opacity = 0:belt5.opacity = 0:belt6.opacity = 50:beltpos = 0
    End Select
    beltpos = beltpos + 1
    if dmdcycle="" then dmdcycle=0
    dmdcycle = dmdcycle+1
    if dmdcycle = 3 Then
    dmdcycle=0
    If UseFlexDMD = 1 then movemovie
    tikiupdate
    if li45.state=1 Then
    credittimer.enabled = true
    Else credittimer.enabled = false
    End If
    End If

  End Sub
dim attractmode


'*********
'Solenoids
'*********

SolCallback(6) = "SolKnocker" 'Unverified
SolCallback(7) = "dtT.SolDropUp"
SolCallback(8) = "dtR2.SolDropUp"
SolCallback(14) = "dtL.SolDropUp"
SolCallback(15) = "bsTrough.SolOut"
SolCallback(17) = "dtR.SolDropUp"
SolCallback(19) = "vpmNudge.SolGameOn"
SolCallback(20) = "SolEMKicker"


Sub SolKnocker(Enabled)
  If enabled Then
    KnockerSolenoid
  End If
End Sub


'EM kicker animation
Dim EMKStep

Sub SolEMKicker(enabled)
    If enabled Then
        EMKStep = 0
        sw32EMK.RotX = 26
        sw32.TimerEnabled = 1
        bsSaucer.ExitSol_On
    SoundSaucerKick 1, sw32

    If GiLights = 1 then launchsound
    If GiLights = 1 then Gi6.state=1
    If GiLights = 1 then Gi7.state=1

    If GiLights = 1 then MusicOn
    If Gi6.state = 1  Then GiOn
    For each x in GiGreen
            x.State = 0
    next

  End If
End Sub

Sub sw32_Timer
    Select Case EMKStep
        Case 1:sw32EMK.RotX = 14
        Case 2:sw32EMK.RotX = 2
        Case 3:sw32EMK.RotX = -20:sw32.TimerEnabled = 0
    End Select
    EMKStep = EMKStep + 1
End Sub



'*******************
' Flipper Subs
'*******************

Const ReflipAngle = 20

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
    If Enabled Then
    LFPress = 1
        LF.fire  'LeftFlipper.RotateToEnd
    LeftFlipper1.RotateToEnd
        If leftflipper.currentangle < leftflipper.endangle + ReflipAngle Then
      RandomSoundReflipUpLeft LeftFlipper
    Else
      SoundFlipperUpAttackLeft LeftFlipper
      RandomSoundFlipperUpLeft LeftFlipper
    End If
    Else
        LFPress = 0
    LeftFlipper.eostorqueangle = EOSA
    LeftFlipper.eostorque = EOST
    LeftFlipper.RotateToStart
        'LeftFlipper1.eostorqueangle = EOSA
    'LeftFlipper1.eostorque = EOST
    LeftFlipper1.RotateToStart
    If LeftFlipper.currentangle < LeftFlipper.startAngle - 5 Then
      RandomSoundFlipperDownLeft LeftFlipper
    End If
        FlipperLeftHitParm = FlipperUpSoundLevel
    End If
End Sub

Sub SolRFlipper(Enabled)
    If Enabled Then
    RFPress = 1
        RF.fire  'RightFlipper.RotateToEnd
        If RightFlipper.currentangle > RightFlipper.endangle - ReflipAngle Then
      RandomSoundReflipUpRight RightFlipper
    Else
      SoundFlipperUpAttackRight RightFlipper
      RandomSoundFlipperUpRight RightFlipper
    End If
    Else
        RFPress = 0
    RightFlipper.eostorqueangle = EOSA
    RightFlipper.eostorque = EOST
    RightFlipper.RotateToStart
    If RightFlipper.currentangle > RightFlipper.startAngle + 5 Then
      RandomSoundFlipperDownRight RightFlipper
    End If
        FlipperRightHitParm = FlipperUpSoundLevel
    End If
End Sub


'*****************
'   Gi Effects
'*****************

Dim OldGiState
OldGiState = -1 'start witht he Gi off

Sub GiON
    For each x in aGiLights
        x.State = 1
    Next
  GiLights=1   'myGi Mod
  StopSound"attract"
  GiMotelR.state=2
  GiMotelB.state=1
  GiMotelG.state=1
  GiMotelTiki.state=2
  vGlow.state=2
  vGlow2.state=2
  vGlow3.state=2
  vGlow4.state=2

  starlight
  'apronstar
  'waves
  'UFOSeq2.Play SeqBlinking, ,6, 250
  'UFOSeq2.Play SeqClockLeftOn, 360, 10
  'UFOSeq2.Play SeqBlinking, ,6, 250

  For each x in aUFOLights
    x.state = 2
next

End Sub


Sub GiFlash
    For each x in aGiLights
        x.Duration 2, 1000,1
   Next

End Sub


Sub GiAlien
For each x in aGiLights
        x.State = 0
Next
For each x in GiGreen
        x.State = 1
  Gi6.state=2
  Gi7.state=2
next
  for each x in ufotwins
  x.state = 2
   Next
End Sub

Sub KillGi
For each x in aGiLights
       x.state=0
next
vGlow.state=0
  vGlow2.state=0
  vGlow3.state=0
  vGlow4.state=0
End Sub


Sub GiRed
For Each x in aGiRed
  x.state=1
Next
End Sub


Sub GiRedOff

For Each x in aGiRed
  x.state=0
Next
End Sub


Sub GiLightning
For each x in aGiLights
        x.State = 0
next
For each x in aFxNight
        x.State = 1

  fx1.state=2
  fx1.BlinkPattern = 0000000010001000010101000110000000000000000010010010001
  fx1.BlinkInterval = 60
  fx2.state=2
  fx2.BlinkPattern = 0000000010001000010101000110000000000000000010010010001
  fx2.BlinkInterval = 50
  fx3.state=2
  fx3.BlinkPattern = 0000000010001000010101000110000000000000000010010010001
  fx3.BlinkInterval = 40

   Next
End Sub


Sub MotelTimer_Timer()
moteltimer.enabled = 0
  GiMotelR.state=2
  GiMotelB.state=1
  GiMotelG.state=1
  GiMotelTiki.state=1
  End Sub


Sub DrainRed
For each x in DrainLighting
        x.State = 1
next
End Sub

Sub DrainNotRed
For each x in DrainLighting
        x.State = 0
next
End Sub


Sub GiOFF
    For each x in aGiLights
        x.State = 0
    Next
  GiLights=0  'myGi Mod
  PlaySound"attract", -1
  nostars
  'nostarsapron
  'nowaves
End Sub

Sub GiEffect(enabled)
    If enabled Then
        For each x in aGiLights
            x.Duration 2, 1000, 1
        Next
      End If
End Sub

Sub GIUpdate
    Dim tmp, obj
    tmp = Getballs
    If UBound(tmp) <> OldGiState Then
        OldGiState = Ubound(tmp)
        If UBound(tmp) = -1 Then
            GiOff
        Else
            GiOn
        End If
    End If

End Sub



'***************************************************
'       JP's VP10 Fading Lamps & Flashers
'       Based on PD's Fading Light System
' SetLamp 0 is Off
' SetLamp 1 is On
' fading for non opacity objects is 4 steps
'***************************************************

Dim LampState(200), FadingLevel(200)
Dim FlashSpeedUp(200), FlashSpeedDown(200), FlashMin(200), FlashMax(200), FlashLevel(200), FlashRepeat(200)

InitLamps()             ' turn off the lights and flashers and reset them to the default parameters
LampTimer.Interval = -1 ' lamp fading speed
LampTimer.Enabled = true

If VRMode = True Then
  Set LampCallback = GetRef("UpdateMultipleLamps")  ' To call VR backglass flasher icons
End If

' Lamp & Flasher Timers

Sub LampTimer_Timer()
    Dim chgLamp, num, chg, ii
    chgLamp = Controller.ChangedLamps
    If Not IsEmpty(chgLamp) Then
        For ii = 0 To UBound(chgLamp)
            LampState(chgLamp(ii, 0) ) = chgLamp(ii, 1)       'keep the real state in an array
            FadingLevel(chgLamp(ii, 0) ) = chgLamp(ii, 1) + 4 'actual fading step
      ' handle flashing DT lights for purpose of sounds effects -apophis
      Select case chgLamp(ii, 0)
        Case 56: if li56.state=1 then dtL_on = 1
        Case 40: if li40.state=1 then dtL_on = 2
        Case 24: if li24.state=1 then dtL_on = 3
        Case 8:  if li8.state=1 then dtL_on = 4
        Case 55: if li55.state=1 then dtL_on = 5
      End Select
        Next
    End If
  If VRMode = True Then
    VRDigitDisplay
  End If
  If VRMode = False Then
    UpdateLeds
  End If
    UpdateLamps
    GIUpdate
    RollingUpdate

End Sub

Sub UpdateLamps()

    'backdrop lights
    If VarHidden Then
        NFadeTm 13, li13, "High Score"
        NFadeT 13, li13a, "To Date"
'        NFadeT 45, li45, "Game Over"
        NFadeT 61, li61, "TILT"
        NFadeT 63, li63, "Match"
        NFadeTm 43, li43a, "Shoot Again"
        If Controller.Lamp(63) Then
            li61a.Text = ""
        Else
            li61a.Text = "Ball in Play"
        End If
    End If
  NFadeL 45, li45
    NFadeL 1, li1
    NFadeL 10, li10
    NFadeL 11, li11
    NFadeL 12, li12
    NFadeL 14, li14
    NFadeL 17, li17
    NFadeL 18, li18
    NFadeL 19, li19
    NFadeL 2, li2
    NFadeL 20, li20
    NFadeL 21, li21
    NFadeL 22, li22
    NFadeL 23, li23
    NFadeLm 23, li23a
    NFadeL 24, li24
    NFadeL 25, li25
    NFadeL 26, li26
    NFadeL 27, li27
    NFadeL 28, li28
    NFadeL 29, li29
    NFadeL 3, li3
    NFadeL 30, li30
    NFadeL 33, li33
    NFadeL 34, li34
    NFadeL 35, li35
    NFadeL 36, li36
    NFadeL 37, li37
    NFadeL 38, li38
    NFadeL 39, li39
    NFadeLm 39, li39a
    NFadeL 4, li4
    NFadeL 40, li40
    NFadeL 41, li41
    NFadeL 42, li42
    NFadeL 43, li43
    NFadeL 44, li44
    NFadeL 46, li46
    NFadeL 49, li49
    NFadeL 5, li5
    NFadeL 50, li50
    NFadeL 51, li51
    NFadeL 52, li52
    NFadeL 53, li53
    NFadeLm 53, li53a
    NFadeL 54, li54
    NFadeL 55, li55
    NFadeL 56, li56
    NFadeL 57, li57
    NFadeL 58, li58
    NFadeL 59, li59
    NFadeL 6, li6
    NFadeL 60, li60
    NFadeL 62, li62
    NFadeL 7, li7
  NFadeLm 7, li7a
    NFadeL 8, li8
    NFadeL 9, li9

End Sub

' div lamp subs

Sub InitLamps()
    Dim x
    For x = 0 to 200
        LampState(x) = 0        ' current light state, independent of the fading level. 0 is off and 1 is on
        FadingLevel(x) = 4      ' used to track the fading state
        FlashSpeedUp(x) = 0.2   ' faster speed when turning on the flasher
        FlashSpeedDown(x) = 0.1 ' slower speed when turning off the flasher
        FlashMax(x) = 1         ' the maximum value when on, usually 1
        FlashMin(x) = 0         ' the minimum value when off, usually 0
        FlashLevel(x) = 0       ' the intensity of the flashers, usually from 0 to 1
        FlashRepeat(x) = 20     ' how many times the flash repeats
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

'Sub NFadeLm(nr, object) ' used for multiple lights
'    Select Case FadingLevel(nr)
'        Case 4:object.state = 0
'        Case 5:object.state = 1
'    End Select
'End Sub

' Mirrors the true on/off state of the lamp (no fade math) 'idig
Sub NFadeLm(nr, object)
    object.State = LampState(nr)
End Sub



'Lights, Ramps & Primitives used as 4 step fading lights
'a,b,c,d are the images used from on to off

Sub FadeObj(nr, object, a, b, c, d)
    Select Case FadingLevel(nr)
        Case 4:object.image = b:FadingLevel(nr) = 6                   'fading to off...
        Case 5:object.image = a:FadingLevel(nr) = 1                   'ON
        Case 6, 7, 8:FadingLevel(nr) = FadingLevel(nr) + 1            'wait
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
            If FlashLevel(nr) <FlashMin(nr) Then
                FlashLevel(nr) = FlashMin(nr)
                FadingLevel(nr) = 0 'completely off
            End if
            Object.IntensityScale = FlashLevel(nr)
        Case 5 ' on
            FlashLevel(nr) = FlashLevel(nr) + FlashSpeedUp(nr)
            If FlashLevel(nr)> FlashMax(nr) Then
                FlashLevel(nr) = FlashMax(nr)
                FadingLevel(nr) = 1 'completely on
            End if
            Object.IntensityScale = FlashLevel(nr)
    End Select
End Sub

Sub Flashm(nr, object) 'multiple flashers, it doesn't change anything, it just follows the main flasher
    Select Case FadingLevel(nr)
        Case 4, 5
            Object.IntensityScale = FlashLevel(nr)
    End Select
End Sub

Sub FlashBlink(nr, object)
    Select Case FadingLevel(nr)
        Case 4 'off
            FlashLevel(nr) = FlashLevel(nr) - FlashSpeedDown(nr)
            If FlashLevel(nr) <FlashMin(nr) Then
                FlashLevel(nr) = FlashMin(nr)
                FadingLevel(nr) = 0 'completely off
            End if
            Object.IntensityScale = FlashLevel(nr)
            If FadingLevel(nr) = 0 AND FlashRepeat(nr) Then 'repeat the flash
                FlashRepeat(nr) = FlashRepeat(nr) -1
                If FlashRepeat(nr) Then FadingLevel(nr) = 5
            End If
        Case 5 ' on
            FlashLevel(nr) = FlashLevel(nr) + FlashSpeedUp(nr)
            If FlashLevel(nr)> FlashMax(nr) Then
                FlashLevel(nr) = FlashMax(nr)
                FadingLevel(nr) = 1 'completely on
            End if
            Object.IntensityScale = FlashLevel(nr)
            If FadingLevel(nr) = 1 AND FlashRepeat(nr) Then FadingLevel(nr) = 4
    End Select
End Sub

' Desktop Objects: Reels & texts (you may also use lights on the desktop)

' Reels

Sub FadeR(nr, object)
    Select Case FadingLevel(nr)
        Case 4:object.SetValue 1:FadingLevel(nr) = 6                   'fading to off...
        Case 5:object.SetValue 0:FadingLevel(nr) = 1                   'ON
        Case 6, 7, 8:FadingLevel(nr) = FadingLevel(nr) + 1             'wait
        Case 9:object.SetValue 2:FadingLevel(nr) = FadingLevel(nr) + 1 'fading...
        Case 10, 11, 12:FadingLevel(nr) = FadingLevel(nr) + 1          'wait
        Case 13:object.SetValue 3:FadingLevel(nr) = 0                  'Off
    End Select
End Sub

Sub FadeRm(nr, object)
    Select Case FadingLevel(nr)
        Case 4:object.SetValue 1
        Case 5:object.SetValue 0
        Case 9:object.SetValue 2
        Case 3:object.SetValue 3
    End Select
End Sub

'Texts

Sub NFadeT(nr, object, message)
    Select Case FadingLevel(nr)
        Case 4:object.Text = "":FadingLevel(nr) = 0
        Case 5:object.Text = message:FadingLevel(nr) = 1
    End Select
End Sub

Sub NFadeTm(nr, object, message)
    Select Case FadingLevel(nr)
        Case 4:object.Text = ""
        Case 5:object.Text = message
    End Select
End Sub

'************************************
'          LEDs Display
'     Based on Scapino's LEDs
'************************************

Dim Digits(32)
Dim Patterns(11)
Dim Patterns2(11)

Patterns(0) = 0     'empty
Patterns(1) = 63    '0
Patterns(2) = 6     '1
Patterns(3) = 91    '2
Patterns(4) = 79    '3
Patterns(5) = 102   '4
Patterns(6) = 109   '5
Patterns(7) = 125   '6
Patterns(8) = 7     '7
Patterns(9) = 127   '8
Patterns(10) = 111  '9

Patterns2(0) = 128  'empty
Patterns2(1) = 191  '0
Patterns2(2) = 134  '1
Patterns2(3) = 219  '2
Patterns2(4) = 207  '3
Patterns2(5) = 230  '4
Patterns2(6) = 237  '5
Patterns2(7) = 253  '6
Patterns2(8) = 135  '7
Patterns2(9) = 255  '8
Patterns2(10) = 239 '9

'Assign 6-digit output to reels
Set Digits(0) = a0
Set Digits(1) = a1
Set Digits(2) = a2
Set Digits(3) = a3
Set Digits(4) = a4
Set Digits(5) = a5
Set Digits(6) = a6

Set Digits(7) = b0
Set Digits(8) = b1
Set Digits(9) = b2
Set Digits(10) = b3
Set Digits(11) = b4
Set Digits(12) = b5
Set Digits(13) = b6

Set Digits(14) = c0
Set Digits(15) = c1
Set Digits(16) = c2
Set Digits(17) = c3
Set Digits(18) = c4
Set Digits(19) = c5
Set Digits(20) = c6

Set Digits(21) = d0
Set Digits(22) = d1
Set Digits(23) = d2
Set Digits(24) = d3
Set Digits(25) = d4
Set Digits(26) = d5
Set Digits(27) = d6

Set Digits(28) = e0
Set Digits(29) = e1
Set Digits(30) = e2
Set Digits(31) = e3

Sub UpdateLeds
    On Error Resume Next
  Dim ChgLED,ii,jj,num,chg,stat,obj
  ChgLED = Controller.ChangedLEDs  (&Hffffffff, &Hffffffff)
  'ChgLED = Controller.ChangedLEDs(&HFF, &HFFFF)
  If Not IsEmpty(ChgLED) Then
    For ii=0 To UBound(chgLED)
      num = chgLED(ii,0): chg=chgLED(ii,1): stat=chgLED(ii,2)
      If UseFlexDMD = 1 then UpdateFlexChar num, stat
      For jj = 0 to 10
                If stat = Patterns(jj) OR stat = Patterns2(jj) then Digits(chgLED(ii, 0) ).SetValue jj
            Next
'     For Each obj In Digits(num)
'       If chg And 1 Then obj.State=stat And 1
'       chg=chg\2:stat=stat\2
'     Next
    Next
    If UseFlexDMD = 1 then FlexDMDUpdate
  End If
End Sub



'Stern Cheetah
'added by Inkochnito
Sub editDips
    Dim vpmDips:Set vpmDips = New cvpmDips
    With vpmDips
        .AddForm 700, 400, "Cheetah - DIP switches"
        .AddFrame 2, 0, 110, "Maximum credits", &H00060000, Array("10 credits", 0, "15 credits", &H00020000, "25 credits", &H00040000, "40 credits", &H00060000) 'dip 18&19
        .AddFrame 2, 76, 110, "High score feature", &H00000020, Array("extra ball", 0, "replay", &H00000020)                                                     'dip 6
        .AddFrame 2, 122, 110, "Balls per game", &H00000040, Array("3 balls", 0, "5 balls", &H00000040)                                                          'dip 7
        .AddFrame 2, 168, 110, "Extra ball award", &H20000000, Array("100K points", 0, "extra ball", &H20000000)                                                 'dip 30
        .AddFrame 125, 0, 110, "High game to date", 49152, Array("points", 0, "1 free game", &H00004000, "2 free games", 32768, "3 free games", 49152)           'dip 15&16
        .AddFrame 125, 76, 110, "Extra ball limit", &H00010000, Array("1 per game", 0, "1 per ball", &H00010000)                                                 'dip 17
        .AddFrame 125, 122, 110, "Special limit", &H00200000, Array("1 per game", 0, "1 per ball", &H00200000)                                                   'dip 22
        .AddFrame 248, 0, 110, "Special award", &HC0000000, Array("no award", 0, "100K points", &H40000000, "free ball", &H80000000, "free game", &HC0000000)    'dip 31&32
        .AddFrame 248, 76, 110, "Add-a-ball maximum", &H00001000, Array("3 balls", 0, "5 balls", &H00001000)                                                     'dip 13
        .AddFrame 248, 122, 110, "5 bank start award", &H10000000, Array("no lite", 0, "20K collect bonus", &H10000000)                                          'dip 29
        .AddChk 125, 170, 110, Array("Match feature", &H00100000)                                                                                                'dip 21
        .AddChk 125, 185, 110, Array("Credits displayed", &H00080000)                                                                                            'dip 20
        .AddChk 125, 200, 110, Array("Bonus memory", &H00800000)                                                                                                 'dip 24
        .AddChk 125, 215, 110, Array("Add-a-ball memory", &H00000010)                                                                                            'dip 5
        .AddChk 248, 170, 110, Array("Background sound", &H00000080)                                                                                             'dip 8
        .AddChk 248, 185, 110, Array("Start with 2X on", &H00002000)                                                                                             'dip 14
        .AddChk 248, 200, 110, Array("3 star bank memory", &H00400000)                                                                                           'dip 23
        .AddLabel 50, 240, 300, 20, "After hitting OK, press F3 to reset game with new settings."
        .ViewDips
    End With
End Sub
Set vpmShowDips = GetRef("editDips")









'******************************************************
'   FLIPPER CORRECTION INITIALIZATION
'******************************************************

dim LF : Set LF = New FlipperPolarity
dim RF : Set RF = New FlipperPolarity

InitPolarity

Sub InitPolarity()
  dim x, a : a = Array(LF, RF)
  for each x in a
    x.AddPoint "Ycoef", 0, RightFlipper.Y-65, 1 'disabled
    x.AddPoint "Ycoef", 1, RightFlipper.Y-11, 1
    x.enabled = True
    x.TimeDelay = 80
  Next

  AddPt "Polarity", 0, 0, 0
  AddPt "Polarity", 1, 0.05, -2.7
  AddPt "Polarity", 2, 0.33, -2.7
  AddPt "Polarity", 3, 0.37, -2.7
  AddPt "Polarity", 4, 0.41, -2.7
  AddPt "Polarity", 5, 0.45, -2.7
  AddPt "Polarity", 6, 0.576,-2.7
  AddPt "Polarity", 7, 0.66, -1.8
  AddPt "Polarity", 8, 0.743, -0.5
  AddPt "Polarity", 9, 0.81, -0.5
  AddPt "Polarity", 10, 0.88, 0

  addpt "Velocity", 0, 0,   1
  addpt "Velocity", 1, 0.16, 1.06
  addpt "Velocity", 2, 0.41,  1.05
  addpt "Velocity", 3, 0.53,  1'0.982
  addpt "Velocity", 4, 0.702, 0.968
  addpt "Velocity", 5, 0.95,  0.968
  addpt "Velocity", 6, 1.03,  0.945

  LF.Object = LeftFlipper
  LF.EndPoint = EndPointLp
  RF.Object = RightFlipper
  RF.EndPoint = EndPointRp
End Sub

Sub TriggerLF_Hit() : LF.Addball activeball : End Sub
Sub TriggerLF_UnHit() : LF.PolarityCorrect activeball : End Sub
Sub TriggerRF_Hit() : RF.Addball activeball : End Sub
Sub TriggerRF_UnHit() : RF.PolarityCorrect activeball : End Sub



'******************************************************
'                        FLIPPER CORRECTION FUNCTIONS
'******************************************************

Sub AddPt(aStr, idx, aX, aY)        'debugger wrapper for adjusting flipper script in-game
        dim a : a = Array(LF, RF)
        dim x : for each x in a
                x.addpoint aStr, idx, aX, aY
        Next
End Sub

Class FlipperPolarity
        Public DebugOn, Enabled
        Private FlipAt        'Timer variable (IE 'flip at 723,530ms...)
        Public TimeDelay        'delay before trigger turns off and polarity is disabled TODO set time!
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

        Public Sub Report(aChooseArray)         'debug, reports all coords in tbPL.text
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
        Private Function FlipperOn() : if gameTime < FlipAt+TimeDelay then FlipperOn = True : End If : End Function        'Timer shutoff for polaritycorrect

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
                                        if ballpos > 0.65 then  Ycoef = LinearEnvelope(BallData(x).Y, YcoefIn, YcoefOut)                                'find safety coefficient 'ycoef' data
                                end if
                        Next

                        If BallPos = 0 Then 'no ball data meaning the ball is entering and exiting pretty close to the same position, use current values.
                                BallPos = PSlope(aBall.x, FlipperStart, 0, FlipperEnd, 1)
                                if ballpos > 0.65 then  Ycoef = LinearEnvelope(aBall.Y, YcoefIn, YcoefOut)                                                'find safety coefficient 'ycoef' data
                        End If

                        'Velocity correction
                        if not IsEmpty(VelocityIn(0) ) then
                                Dim VelCoef
         :                         VelCoef = LinearEnvelope(BallPos, VelocityIn, VelocityOut)

                                if partialflipcoef < 1 then VelCoef = PSlope(partialflipcoef, 0, 1, 1, VelCoef)

                                if Enabled then aBall.Velx = aBall.Velx*VelCoef
                                if Enabled then aBall.Vely = aBall.Vely*VelCoef
                        End If

                        'Polarity Correction (optional now)
                        if not IsEmpty(PolarityIn(0) ) then
                                If StartPoint > EndPoint then LR = -1        'Reverse polarity if left flipper
                                dim AddX : AddX = LinearEnvelope(BallPos, PolarityIn, PolarityOut) * LR

                                if Enabled then aBall.VelX = aBall.VelX + 1 * (AddX*ycoef*PartialFlipcoef)
                                'playsound "fx_knocker"
                        End If
                End If
                RemoveBall aBall
        End Sub
End Class

'******************************************************
'                FLIPPER POLARITY AND RUBBER DAMPENER
'                        SUPPORTING FUNCTIONS
'******************************************************

' Used for flipper correction and rubber dampeners
Sub ShuffleArray(ByRef aArray, byVal offset) 'shuffle 1d array
        dim x, aCount : aCount = 0
        redim a(uBound(aArray) )
        for x = 0 to uBound(aArray)        'Shuffle objects in a temp array
                if not IsEmpty(aArray(x) ) Then
                        if IsObject(aArray(x)) then
                                Set a(aCount) = aArray(x)
                        Else
                                a(aCount) = aArray(x)
                        End If
                        aCount = aCount + 1
                End If
        Next
        if offset < 0 then offset = 0
        redim aArray(aCount-1+offset)        'Resize original array
        for x = 0 to aCount-1                'set objects back into original array
                if IsObject(a(x)) then
                        Set aArray(x) = a(x)
                Else
                        aArray(x) = a(x)
                End If
        Next
End Sub

' Used for flipper correction and rubber dampeners
Sub ShuffleArrays(aArray1, aArray2, offset)
        ShuffleArray aArray1, offset
        ShuffleArray aArray2, offset
End Sub

' Used for flipper correction, rubber dampeners, and drop targets
Function BallSpeed(ball) 'Calculates the ball speed
    BallSpeed = SQR(ball.VelX^2 + ball.VelY^2 + ball.VelZ^2)
End Function

' Used for flipper correction and rubber dampeners
Function PSlope(Input, X1, Y1, X2, Y2)        'Set up line via two points, no clamping. Input X, output Y
        dim x, y, b, m : x = input : m = (Y2 - Y1) / (X2 - X1) : b = Y2 - m*X2
        Y = M*x+b
        PSlope = Y
End Function

' Used for flipper correction
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

' Used for flipper correction and rubber dampeners
Function LinearEnvelope(xInput, xKeyFrame, yLvl)
        dim y 'Y output
        dim L 'Line
        dim ii : for ii = 1 to uBound(xKeyFrame)        'find active line
                if xInput <= xKeyFrame(ii) then L = ii : exit for : end if
        Next
        if xInput > xKeyFrame(uBound(xKeyFrame) ) then L = uBound(xKeyFrame)        'catch line overrun
        Y = pSlope(xInput, xKeyFrame(L-1), yLvl(L-1), xKeyFrame(L), yLvl(L) )

        if xInput <= xKeyFrame(lBound(xKeyFrame) ) then Y = yLvl(lBound(xKeyFrame) )         'Clamp lower
        if xInput >= xKeyFrame(uBound(xKeyFrame) ) then Y = yLvl(uBound(xKeyFrame) )        'Clamp upper

        LinearEnvelope = Y
End Function





'******************************************************
'                        FLIPPER TRICKS
'******************************************************

RightFlipper.timerinterval=1
Rightflipper.timerenabled=True

sub RightFlipper_timer()
        FlipperTricks LeftFlipper, LFPress, LFCount, LFEndAngle, LFState
        FlipperTricks RightFlipper, RFPress, RFCount, RFEndAngle, RFState
        FlipperNudge RightFlipper, RFEndAngle, RFEOSNudge, LeftFlipper, LFEndAngle
        FlipperNudge LeftFlipper, LFEndAngle, LFEOSNudge,  RightFlipper, RFEndAngle
end sub

Dim LFEOSNudge, RFEOSNudge

Sub FlipperNudge(Flipper1, Endangle1, EOSNudge1, Flipper2, EndAngle2)
        Dim BOT, b

        If Flipper1.currentangle = Endangle1 and EOSNudge1 <> 1 Then
                EOSNudge1 = 1
                'debug.print Flipper1.currentangle &" = "& Endangle1 &"--"& Flipper2.currentangle &" = "& EndAngle2
                If Flipper2.currentangle = EndAngle2 Then
                        BOT = GetBalls
                        For b = 0 to Ubound(BOT)
                                If FlipperTrigger(BOT(b).x, BOT(b).y, Flipper1) Then
                                        'Debug.Print "ball in flip1. exit"
                                         exit Sub
                                end If
                        Next
                        For b = 0 to Ubound(BOT)
                                If FlipperTrigger(BOT(b).x, BOT(b).y, Flipper2) Then
                                        BOT(b).velx = BOT(b).velx / 1.3
                                        BOT(b).vely = BOT(b).vely - 0.5
                                end If
                        Next
                End If
        Else
                If Flipper1.currentangle <> EndAngle1 then
                        EOSNudge1 = 0
                end if
        End If
End Sub

'*****************
' Maths
'*****************
Const PI = 3.1415927

Function dSin(degrees)
        dsin = sin(degrees * Pi/180)
End Function

Function dCos(degrees)
        dcos = cos(degrees * Pi/180)
End Function

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

' Used for drop targets and stand up targets
Function Atn2(dy, dx)
        dim pi
        pi = 4*Atn(1)

        If dx > 0 Then
                Atn2 = Atn(dy / dx)
        ElseIf dx < 0 Then
                If dy = 0 Then
                        Atn2 = pi
                Else
                        Atn2 = Sgn(dy) * (pi - Atn(Abs(dy / dx)))
                end if
        ElseIf dx = 0 Then
                if dy = 0 Then
                        Atn2 = 0
                else
                        Atn2 = Sgn(dy) * pi / 2
                end if
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
Const EOSTnew = 0.8
Const EOSAnew = 1
Const EOSRampup = 0
Dim SOSRampup
Select Case FlipperCoilRampupMode
  Case 0:
    SOSRampup = 2.5
  Case 1:
    SOSRampup = 6
  Case 2:
    SOSRampup = 8.5
End Select
Const LiveCatch = 16
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
        Dir = Flipper.startangle/Abs(Flipper.startangle)        '-1 for Right Flipper

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
        Dim LiveCatchBounce                                                                                                                        'If live catch is not perfect, it won't freeze ball totally
        Dim CatchTime : CatchTime = GameTime - FCount

        if CatchTime <= LiveCatch and parm > 6 and ABS(Flipper.x - ball.x) > LiveDistanceMin and ABS(Flipper.x - ball.x) < LiveDistanceMax Then
                if CatchTime <= LiveCatch*0.5 Then                                                'Perfect catch only when catch time happens in the beginning of the window
                        LiveCatchBounce = 0
                else
                        LiveCatchBounce = Abs((LiveCatch/2) - CatchTime)        'Partial catch when catch happens a bit late
                end If

                If LiveCatchBounce = 0 and ball.velx * Dir > 0 Then ball.velx = 0
                ball.vely = LiveCatchBounce * (32 / LiveCatch) ' Multiplier for inaccuracy bounce
                ball.angmomx= 0
                ball.angmomy= 0
                ball.angmomz= 0
        End If
End Sub

Sub LeftFlipper_Collide(parm)
  CheckLiveCatch Activeball, LeftFlipper, LFCount, parm
  LeftFlipperCollide parm
  if RubberizerEnabled <> 0 then Rubberizer(parm)
End Sub

Sub RightFlipper_Collide(parm)
  CheckLiveCatch Activeball, RightFlipper, RFCount, parm
  RightFlipperCollide parm
  if RubberizerEnabled <> 0 then Rubberizer(parm)
End Sub

sub Rubberizer(parm)
  if parm < 10 And parm > 2 And Abs(activeball.angmomz) < 10 then
    'debug.print "parm: " & parm & " momz: " & activeball.angmomz &" vely: "& activeball.vely
    activeball.angmomz = activeball.angmomz * 1.2
    activeball.vely = activeball.vely * 1.2
    'debug.print ">> newmomz: " & activeball.angmomz&" newvely: "& activeball.vely
  Elseif parm <= 2 and parm > 0.2 Then
    'debug.print "* parm: " & parm & " momz: " & activeball.angmomz &" vely: "& activeball.vely
    activeball.angmomz = activeball.angmomz * -1.1
    activeball.vely = activeball.vely * 1.4
    'debug.print "**** >> newmomz: " & activeball.angmomz&" newvely: "& activeball.vely
  end if
end sub


'****************************************************************************
'PHYSICS DAMPENERS

'These are data mined bounce curves,
'dialed in with the in-game elasticity as much as possible to prevent angle / spin issues.
'Requires tracking ballspeed to calculate COR


Sub dPosts_Hit(idx)
        RubbersD.dampen Activeball
    TargetBouncer Activeball, 1
End Sub

Sub dSleeves_Hit(idx)
        SleevesD.Dampen Activeball
    TargetBouncer Activeball, 0.7
End Sub


dim RubbersD : Set RubbersD = new Dampener        'frubber
RubbersD.name = "Rubbers"
RubbersD.debugOn = False        'shows info in textbox "TBPout"
RubbersD.Print = False        'debug, reports in debugger (in vel, out cor)
'cor bounce curve (linear)
'for best results, try to match in-game velocity as closely as possible to the desired curve
'RubbersD.addpoint 0, 0, 0.935        'point# (keep sequential), ballspeed, CoR (elasticity)
RubbersD.addpoint 0, 0, 0.96        'point# (keep sequential), ballspeed, CoR (elasticity)
RubbersD.addpoint 1, 3.77, 0.96
RubbersD.addpoint 2, 5.76, 0.967        'dont take this as gospel. if you can data mine rubber elasticitiy, please help!
RubbersD.addpoint 3, 15.84, 0.874
RubbersD.addpoint 4, 56, 0.64        'there's clamping so interpolate up to 56 at least

dim SleevesD : Set SleevesD = new Dampener        'this is just rubber but cut down to 85%...
SleevesD.name = "Sleeves"
SleevesD.debugOn = False        'shows info in textbox "TBPout"
SleevesD.Print = False        'debug, reports in debugger (in vel, out cor)
SleevesD.CopyCoef RubbersD, 0.85

Class Dampener
        Public Print, debugOn 'tbpOut.text
        public name, Threshold         'Minimum threshold. Useful for Flippers, which don't have a hit threshold.
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
                RealCOR = BallSpeed(aBall) / cor.ballvel(aBall.id)
                coef = desiredcor / realcor
                if debugOn then str = name & " in vel:" & round(cor.ballvel(aBall.id),2 ) & vbnewline & "desired cor: " & round(desiredcor,4) & vbnewline & _
                "actual cor: " & round(realCOR,4) & vbnewline & "ballspeed coef: " & round(coef, 3) & vbnewline
                if Print then debug.print Round(cor.ballvel(aBall.id),2) & ", " & round(desiredcor,3)

' Thalamus - patched :                 aBall.velx = aBall.velx * coef : aBall.vely = aBall.vely * coef
    aBall.velx = aBall.velx * coef : aBall.vely = aBall.vely * coef : aBall.velz = aBall.velz * coef
                if debugOn then TBPout.text = str
        End Sub

        Public Sub CopyCoef(aObj, aCoef) 'alternative addpoints, copy with coef
                dim x : for x = 0 to uBound(aObj.ModIn)
                        addpoint x, aObj.ModIn(x), aObj.ModOut(x)*aCoef
                Next
        End Sub


        Public Sub Report()         'debug, reports all coords in tbPL.text
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


'******************************************************
'iaakki - TargetBouncer for targets and posts
'******************************************************
Dim zMultiplier

sub TargetBouncer(aBall,defvalue)
  if TargetBouncerEnabled <> 0 and aball.z < 30 then
    'debug.print "velz: " & activeball.velz
    Select Case Int(Rnd * 4) + 1
      Case 1: zMultiplier = defvalue+1.1
      Case 2: zMultiplier = defvalue+1.05
      Case 3: zMultiplier = defvalue+0.7
      Case 4: zMultiplier = defvalue+0.3
    End Select
    aBall.velz = aBall.velz * zMultiplier * TargetBouncerFactor
    'debug.print "----> velz: " & activeball.velz
  end if
end sub



'////////////////////////////  MECHANICAL SOUNDS  ///////////////////////////
'//  This part in the script is an entire block that is dedicated to the physics sound system.
'//  Various scripts and sounds that may be pretty generic and could suit other WPC systems, but the most are tailored specifically for this table.

'///////////////////////////////  SOUNDS PARAMETERS  //////////////////////////////
Dim GlobalSoundLevel, CoinSoundLevel, PlungerReleaseSoundLevel, PlungerPullSoundLevel, NudgeLeftSoundLevel, SpinSoundLevel
Dim NudgeRightSoundLevel, NudgeCenterSoundLevel, StartButtonSoundLevel, RollingSoundFactor

SpinSoundLevel = 2    ' Spinner Volume - higher value, higher sound
CoinSoundLevel = 1                                                                                                                'volume level; range [0, 1]
NudgeLeftSoundLevel = 1                                                                                                        'volume level; range [0, 1]
NudgeRightSoundLevel = 1                                                                                                'volume level; range [0, 1]
NudgeCenterSoundLevel = 1                                                                                                'volume level; range [0, 1]
StartButtonSoundLevel = 0.1                                                                                                'volume level; range [0, 1]
PlungerReleaseSoundLevel = 0.8 '1 wjr                                                                                        'volume level; range [0, 1]
PlungerPullSoundLevel = 1                                                                                                'volume level; range [0, 1]
RollingSoundFactor = 1.1/5

'///////////////////////-----Solenoids, Kickers and Flash Relays-----///////////////////////
Dim FlipperUpAttackMinimumSoundLevel, FlipperUpAttackMaximumSoundLevel, FlipperUpAttackLeftSoundLevel, FlipperUpAttackRightSoundLevel
Dim FlipperUpSoundLevel, FlipperDownSoundLevel, FlipperLeftHitParm, FlipperRightHitParm
Dim SlingshotSoundLevel, BumperSoundFactor, KnockerSoundLevel

FlipperUpAttackMinimumSoundLevel = 0.010                                                           'volume level; range [0, 1]
FlipperUpAttackMaximumSoundLevel = 0.635                                                                'volume level; range [0, 1]
FlipperUpSoundLevel = 1.0                                                                        'volume level; range [0, 1]
FlipperDownSoundLevel = 0.45                                                                      'volume level; range [0, 1]
FlipperLeftHitParm = FlipperUpSoundLevel                                                                'sound helper; not configurable
FlipperRightHitParm = FlipperUpSoundLevel                                                                'sound helper; not configurable
SlingshotSoundLevel = 0.95                                                                                                'volume level; range [0, 1]
BumperSoundFactor = 4.25                                                                                                'volume multiplier; must not be zero
KnockerSoundLevel = 1                                                                                                         'volume level; range [0, 1]

'///////////////////////-----Ball Drops, Bumps and Collisions-----///////////////////////
Dim RubberStrongSoundFactor, RubberWeakSoundFactor, RubberFlipperSoundFactor,BallWithBallCollisionSoundFactor
Dim BallBouncePlayfieldSoftFactor, BallBouncePlayfieldHardFactor, PlasticRampDropToPlayfieldSoundLevel, WireRampDropToPlayfieldSoundLevel, DelayedBallDropOnPlayfieldSoundLevel
Dim WallImpactSoundFactor, MetalImpactSoundFactor, SubwaySoundLevel, SubwayEntrySoundLevel, ScoopEntrySoundLevel
Dim SaucerLockSoundLevel, SaucerKickSoundLevel

BallWithBallCollisionSoundFactor = 3.2                                                                        'volume multiplier; must not be zero
RubberStrongSoundFactor = 0.055/5                                                                                        'volume multiplier; must not be zero
RubberWeakSoundFactor = 0.075/5                                                                                        'volume multiplier; must not be zero
RubberFlipperSoundFactor = 0.075/5                                                                                'volume multiplier; must not be zero
BallBouncePlayfieldSoftFactor = 0.025                                                                        'volume multiplier; must not be zero
BallBouncePlayfieldHardFactor = 0.025                                                                        'volume multiplier; must not be zero
DelayedBallDropOnPlayfieldSoundLevel = 0.8                                                                        'volume level; range [0, 1]
WallImpactSoundFactor = 0.075                                                                                        'volume multiplier; must not be zero
MetalImpactSoundFactor = 0.075/1000
SaucerLockSoundLevel = 0.8
SaucerKickSoundLevel = 0.8

'///////////////////////-----Gates, Spinners, Rollovers and Targets-----///////////////////////

Dim GateSoundLevel, TargetSoundFactor, SpinnerSoundLevel, RolloverSoundLevel, DTSoundLevel

GateSoundLevel = 0.5/5                                                                                                        'volume level; range [0, 1]
TargetSoundFactor = 0.0025 * 10                                                                                        'volume multiplier; must not be zero
DTSoundLevel = 0.25                                                                                                                'volume multiplier; must not be zero
RolloverSoundLevel = 0.25                                                                      'volume level; range [0, 1]

'///////////////////////-----Ball Release, Guides and Drain-----///////////////////////
Dim DrainSoundLevel, BallReleaseSoundLevel, BottomArchBallGuideSoundFactor, FlipperBallGuideSoundFactor

DrainSoundLevel = 0.8                                                                                                                'volume level; range [0, 1]
BallReleaseSoundLevel = 1                                                                                                'volume level; range [0, 1]
BottomArchBallGuideSoundFactor = 0.2                                                                        'volume multiplier; must not be zero
FlipperBallGuideSoundFactor = 0.015                                                                                'volume multiplier; must not be zero

'///////////////////////-----Loops and Lanes-----///////////////////////
Dim ArchSoundFactor
ArchSoundFactor = 0.025/5                                                                                                       'volume multiplier; must not be zero


'/////////////////////////////  SOUND PLAYBACK FUNCTIONS  ////////////////////////////
'/////////////////////////////  POSITIONAL SOUND PLAYBACK METHODS  ////////////////////////////
' Positional sound playback methods will play a sound, depending on the X,Y position of the table element or depending on ActiveBall object position
' These are similar subroutines that are less complicated to use (e.g. simply use standard parameters for the PlaySound call)
' For surround setup - positional sound playback functions will fade between front and rear surround channels and pan between left and right channels
' For stereo setup - positional sound playback functions will only pan between left and right channels
' For mono setup - positional sound playback functions will not pan between left and right channels and will not fade between front and rear channels

' PlaySound full syntax - PlaySound(string, int loopcount, float volume, float pan, float randompitch, int pitch, bool useexisting, bool restart, float front_rear_fade)
' Note - These functions will not work (currently) for walls/slingshots as these do not feature a simple, single X,Y position
Sub PlaySoundAtLevelStatic(playsoundparams, aVol, tableobj)
    PlaySound playsoundparams, 0, aVol * VolumeDial, AudioPan(tableobj), 0, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelExistingStatic(playsoundparams, aVol, tableobj)
    PlaySound playsoundparams, 0, aVol * VolumeDial, AudioPan(tableobj), 0, 0, 1, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelStaticLoop(playsoundparams, aVol, tableobj)
    PlaySound playsoundparams, -1, aVol * VolumeDial, AudioPan(tableobj), 0, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelStaticRandomPitch(playsoundparams, aVol, randomPitch, tableobj)
    PlaySound playsoundparams, 0, aVol * VolumeDial, AudioPan(tableobj), randomPitch, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelActiveBall(playsoundparams, aVol)
        PlaySound playsoundparams, 0, aVol * VolumeDial, AudioPan(ActiveBall), 0, 0, 0, 0, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtLevelExistingActiveBall(playsoundparams, aVol)
        PlaySound playsoundparams, 0, aVol * VolumeDial, AudioPan(ActiveBall), 0, 0, 1, 0, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtLeveTimerActiveBall(playsoundparams, aVol, ballvariable)
        PlaySound playsoundparams, 0, aVol * VolumeDial, AudioPan(ballvariable), 0, 0, 0, 0, AudioFade(ballvariable)
End Sub

Sub PlaySoundAtLevelTimerExistingActiveBall(playsoundparams, aVol, ballvariable)
        PlaySound playsoundparams, 0, aVol * VolumeDial, AudioPan(ballvariable), 0, 0, 1, 0, AudioFade(ballvariable)
End Sub

Sub PlaySoundAtLevelRoll(playsoundparams, aVol, pitch)
    PlaySound playsoundparams, -1, aVol * VolumeDial, AudioPan(tableobj), randomPitch, 0, 0, 0, AudioFade(tableobj)
End Sub

' Previous Positional Sound Subs

Sub PlaySoundAt(soundname, tableobj)
    PlaySound soundname, 1, 1 * VolumeDial, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

Sub PlaySoundAtVol(soundname, tableobj, aVol)
    PlaySound soundname, 1, aVol * VolumeDial, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

Sub PlaySoundAtBall(soundname)
    PlaySoundAt soundname, ActiveBall
End Sub

Sub PlaySoundAtBallVol (Soundname, aVol)
        Playsound soundname, 1,aVol * VolumeDial, AudioPan(ActiveBall), 0,0,0, 1, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtBallVolM (Soundname, aVol)
        Playsound soundname, 1,aVol * VolumeDial, AudioPan(ActiveBall), 0,0,0, 0, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtVolLoops(sound, tableobj, Vol, Loops)
        PlaySound sound, Loops, Vol * VolumeDial, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub


' *********************************************************************
'                     Fleep  Supporting Ball & Sound Functions
' *********************************************************************

Dim tablewidth, tableheight : tablewidth = table1.width : tableheight = table1.height

Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "table1" is the name of the table
        Dim tmp
    tmp = tableobj.y * 2 / tableheight-1
    If tmp > 0 Then
                AudioFade = Csng(tmp ^10)
    Else
        AudioFade = Csng(-((- tmp) ^10) )
    End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "table1" is the name of the table
    Dim tmp
    tmp = tableobj.x * 2 / tablewidth-1
    If tmp > 0 Then
        AudioPan = Csng(tmp ^10)
    Else
        AudioPan = Csng(-((- tmp) ^10) )
    End If
End Function

Function Vol(ball) ' Calculates the volume of the sound based on the ball speed
        Vol = Csng(BallVel(ball) ^2)
End Function

Function Volz(ball) ' Calculates the volume of the sound based on the ball speed
        Volz = Csng((ball.velz) ^2)
End Function

Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
    Pitch = BallVel(ball) * 20
End Function

Function BallVel(ball) 'Calculates the ball speed
    BallVel = INT(SQR((ball.VelX ^2) + (ball.VelY ^2) ) )
End Function

Function VolPlayfieldRoll(ball) ' Calculates the roll volume of the sound based on the ball speed
        VolPlayfieldRoll = RollingSoundFactor * 0.0005 * Csng(BallVel(ball) ^3)
End Function

Function PitchPlayfieldRoll(ball) ' Calculates the roll pitch of the sound based on the ball speed
    PitchPlayfieldRoll = BallVel(ball) ^2 * 15
End Function

Function RndInt(min, max)
    RndInt = Int(Rnd() * (max-min + 1) + min)' Sets a random number integer between min and max
End Function

Function RndNum(min, max)
    RndNum = Rnd() * (max-min) + min' Sets a random number between min and max
End Function

'/////////////////////////////  GENERAL SOUND SUBROUTINES  ////////////////////////////
Sub SoundStartButton()
        PlaySound ("Start_Button"), 0, StartButtonSoundLevel, 0, 0.25
End Sub

Sub SoundNudgeLeft()
        PlaySound ("Nudge_" & Int(Rnd*2)+1), 0, NudgeLeftSoundLevel * VolumeDial, -0.1, 0.25
End Sub

Sub SoundNudgeRight()
        PlaySound ("Nudge_" & Int(Rnd*2)+1), 0, NudgeRightSoundLevel * VolumeDial, 0.1, 0.25
End Sub

Sub SoundNudgeCenter()
        PlaySound ("Nudge_" & Int(Rnd*2)+1), 0, NudgeCenterSoundLevel * VolumeDial, 0, 0.25
End Sub


Sub SoundPlungerPull()
        PlaySoundAtLevelStatic ("Plunger_Pull_1"), PlungerPullSoundLevel, Plunger
End Sub

Sub SoundPlungerReleaseBall()
        PlaySoundAtLevelStatic ("Plunger_Release_Ball"), PlungerReleaseSoundLevel, Plunger
End Sub

Sub SoundPlungerReleaseNoBall()
        PlaySoundAtLevelStatic ("Plunger_Release_No_Ball"), PlungerReleaseSoundLevel, Plunger
End Sub


'/////////////////////////////  KNOCKER SOLENOID  ////////////////////////////
Sub KnockerSolenoid()
        PlaySoundAtLevelStatic SoundFX("Knocker_1",DOFKnocker), KnockerSoundLevel, KnockerPosition
End Sub

'/////////////////////////////  DRAIN SOUNDS  ////////////////////////////
Sub RandomSoundDrain(drainswitch)
        PlaySoundAtLevelStatic ("Drain_" & Int(Rnd*11)+1), DrainSoundLevel, drainswitch
End Sub

'/////////////////////////////  TROUGH BALL RELEASE SOLENOID SOUNDS  ////////////////////////////

Sub RandomSoundBallRelease(drainswitch)
        PlaySoundAtLevelStatic SoundFX("BallRelease" & Int(Rnd*7)+1,DOFContactors), BallReleaseSoundLevel, drainswitch
End Sub

'/////////////////////////////  SLINGSHOT SOLENOID SOUNDS  ////////////////////////////
Sub RandomSoundSlingshotLeft(sling)
        PlaySoundAtLevelStatic SoundFX("Sling_L" & Int(Rnd*10)+1,DOFContactors), SlingshotSoundLevel, Sling
End Sub

Sub RandomSoundSlingshotRight(sling)
        PlaySoundAtLevelStatic SoundFX("Sling_R" & Int(Rnd*8)+1,DOFContactors), SlingshotSoundLevel, Sling
End Sub

'/////////////////////////////  BUMPER SOLENOID SOUNDS  ////////////////////////////
Sub RandomSoundBumperTop(Bump)
        PlaySoundAtLevelStatic SoundFX("Bumpers_Top_" & Int(Rnd*5)+1,DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
End Sub

Sub RandomSoundBumperMiddle(Bump)
        PlaySoundAtLevelStatic SoundFX("Bumpers_Middle_" & Int(Rnd*5)+1,DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
End Sub

Sub RandomSoundBumperBottom(Bump)
        PlaySoundAtLevelStatic SoundFX("Bumpers_Bottom_" & Int(Rnd*5)+1,DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
End Sub

'/////////////////////////////  FLIPPER BATS SOUND SUBROUTINES  ////////////////////////////
'/////////////////////////////  FLIPPER BATS SOLENOID ATTACK SOUND  ////////////////////////////
Sub SoundFlipperUpAttackLeft(flipper)
        FlipperUpAttackLeftSoundLevel = RndNum(FlipperUpAttackMinimumSoundLevel, FlipperUpAttackMaximumSoundLevel)
        PlaySoundAtLevelStatic ("Flipper_Attack-L01"), FlipperUpAttackLeftSoundLevel, flipper
End Sub

Sub SoundFlipperUpAttackRight(flipper)
        FlipperUpAttackRightSoundLevel = RndNum(FlipperUpAttackMinimumSoundLevel, FlipperUpAttackMaximumSoundLevel)
                PlaySoundAtLevelStatic ("Flipper_Attack-R01"), FlipperUpAttackLeftSoundLevel, flipper
End Sub

'/////////////////////////////  FLIPPER BATS SOLENOID CORE SOUND  ////////////////////////////
Sub RandomSoundFlipperUpLeft(flipper)
        PlaySoundAtLevelStatic SoundFX("Flipper_L0" & Int(Rnd*9)+1,DOFFlippers), FlipperLeftHitParm, Flipper
End Sub

Sub RandomSoundFlipperUpRight(flipper)
        PlaySoundAtLevelStatic SoundFX("Flipper_R0" & Int(Rnd*9)+1,DOFFlippers), FlipperRightHitParm, Flipper
End Sub

Sub RandomSoundReflipUpLeft(flipper)
        PlaySoundAtLevelStatic SoundFX("Flipper_ReFlip_L0" & Int(Rnd*3)+1,DOFFlippers), (RndNum(0.8, 1))*FlipperUpSoundLevel, Flipper
End Sub

Sub RandomSoundReflipUpRight(flipper)
        PlaySoundAtLevelStatic SoundFX("Flipper_ReFlip_R0" & Int(Rnd*3)+1,DOFFlippers), (RndNum(0.8, 1))*FlipperUpSoundLevel, Flipper
End Sub

Sub RandomSoundFlipperDownLeft(flipper)
        PlaySoundAtLevelStatic SoundFX("Flipper_Left_Down_" & Int(Rnd*7)+1,DOFFlippers), FlipperDownSoundLevel, Flipper
End Sub

Sub RandomSoundFlipperDownRight(flipper)
        PlaySoundAtLevelStatic SoundFX("Flipper_Right_Down_" & Int(Rnd*8)+1,DOFFlippers), FlipperDownSoundLevel, Flipper
End Sub

'/////////////////////////////  FLIPPER BATS BALL COLLIDE SOUND  ////////////////////////////

Sub LeftFlipperCollide(parm)
        FlipperLeftHitParm = parm/10
        If FlipperLeftHitParm > 1 Then
                FlipperLeftHitParm = 1
        End If
        FlipperLeftHitParm = FlipperUpSoundLevel * FlipperLeftHitParm
        RandomSoundRubberFlipper(parm)
End Sub

Sub RightFlipperCollide(parm)
        FlipperRightHitParm = parm/10
        If FlipperRightHitParm > 1 Then
                FlipperRightHitParm = 1
        End If
        FlipperRightHitParm = FlipperUpSoundLevel * FlipperRightHitParm
         RandomSoundRubberFlipper(parm)
End Sub

Sub RandomSoundRubberFlipper(parm)
        PlaySoundAtLevelActiveBall ("Flipper_Rubber_" & Int(Rnd*7)+1), parm  * RubberFlipperSoundFactor
End Sub

'/////////////////////////////  ROLLOVER SOUNDS  ////////////////////////////
Sub RandomSoundRollover()
        PlaySoundAtLevelActiveBall ("Rollover_" & Int(Rnd*4)+1), RolloverSoundLevel
End Sub

Sub Rollovers_Hit(idx)
        RandomSoundRollover
End Sub

'/////////////////////////////  VARIOUS PLAYFIELD SOUND SUBROUTINES  ////////////////////////////
'/////////////////////////////  RUBBERS AND POSTS  ////////////////////////////
'/////////////////////////////  RUBBERS - EVENTS  ////////////////////////////
Sub Rubbers_Hit(idx)
         dim finalspeed
          finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
         If finalspeed > 5 then
                 RandomSoundRubberStrong 1
        End if
        If finalspeed <= 5 then
                 RandomSoundRubberWeak()
         End If
End Sub

'/////////////////////////////  RUBBERS AND POSTS - STRONG IMPACTS  ////////////////////////////
Sub RandomSoundRubberStrong(voladj)
        Select Case Int(Rnd*10)+1
                Case 1 : PlaySoundAtLevelActiveBall ("Rubber_Strong_1"), Vol(ActiveBall) * RubberStrongSoundFactor*voladj
                Case 2 : PlaySoundAtLevelActiveBall ("Rubber_Strong_2"), Vol(ActiveBall) * RubberStrongSoundFactor*voladj
                Case 3 : PlaySoundAtLevelActiveBall ("Rubber_Strong_3"), Vol(ActiveBall) * RubberStrongSoundFactor*voladj
                Case 4 : PlaySoundAtLevelActiveBall ("Rubber_Strong_4"), Vol(ActiveBall) * RubberStrongSoundFactor*voladj
                Case 5 : PlaySoundAtLevelActiveBall ("Rubber_Strong_5"), Vol(ActiveBall) * RubberStrongSoundFactor*voladj
                Case 6 : PlaySoundAtLevelActiveBall ("Rubber_Strong_6"), Vol(ActiveBall) * RubberStrongSoundFactor*voladj
                Case 7 : PlaySoundAtLevelActiveBall ("Rubber_Strong_7"), Vol(ActiveBall) * RubberStrongSoundFactor*voladj
                Case 8 : PlaySoundAtLevelActiveBall ("Rubber_Strong_8"), Vol(ActiveBall) * RubberStrongSoundFactor*voladj
                Case 9 : PlaySoundAtLevelActiveBall ("Rubber_Strong_9"), Vol(ActiveBall) * RubberStrongSoundFactor*voladj
                Case 10 : PlaySoundAtLevelActiveBall ("Rubber_1_Hard"), Vol(ActiveBall) * RubberStrongSoundFactor * 0.6*voladj
        End Select
End Sub

'/////////////////////////////  RUBBERS AND POSTS - WEAK IMPACTS  ////////////////////////////
Sub RandomSoundRubberWeak()
        PlaySoundAtLevelActiveBall ("Rubber_" & Int(Rnd*9)+1), Vol(ActiveBall) * RubberWeakSoundFactor
End Sub

'/////////////////////////////  WALL IMPACTS  ////////////////////////////
Sub Walls_Hit(idx)
         dim finalspeed
          finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
         If finalspeed > 5 then
                 RandomSoundRubberStrong 1
        End if
        If finalspeed <= 5 then
                 RandomSoundRubberWeak()
         End If
End Sub

Sub RandomSoundWall()
         dim finalspeed
          finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
         If finalspeed > 16 then
                Select Case Int(Rnd*5)+1
                        Case 1 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_1"), Vol(ActiveBall) * WallImpactSoundFactor
                        Case 2 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_2"), Vol(ActiveBall) * WallImpactSoundFactor
                        Case 3 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_5"), Vol(ActiveBall) * WallImpactSoundFactor
                        Case 4 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_7"), Vol(ActiveBall) * WallImpactSoundFactor
                        Case 5 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_9"), Vol(ActiveBall) * WallImpactSoundFactor
                End Select
        End if
        If finalspeed >= 6 AND finalspeed <= 16 then
                Select Case Int(Rnd*4)+1
                        Case 1 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_3"), Vol(ActiveBall) * WallImpactSoundFactor
                        Case 2 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_4"), Vol(ActiveBall) * WallImpactSoundFactor
                        Case 3 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_6"), Vol(ActiveBall) * WallImpactSoundFactor
                        Case 4 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_8"), Vol(ActiveBall) * WallImpactSoundFactor
                End Select
         End If
        If finalspeed < 6 Then
                Select Case Int(Rnd*3)+1
                        Case 1 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_4"), Vol(ActiveBall) * WallImpactSoundFactor
                        Case 2 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_6"), Vol(ActiveBall) * WallImpactSoundFactor
                        Case 3 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_8"), Vol(ActiveBall) * WallImpactSoundFactor
                End Select
        End if
End Sub

'/////////////////////////////  METAL TOUCH SOUNDS  ////////////////////////////
Sub RandomSoundMetal()
        PlaySoundAtLevelActiveBall ("Metal_Touch_" & Int(Rnd*13)+1), Vol(ActiveBall) * MetalImpactSoundFactor
End Sub

'/////////////////////////////  METAL - EVENTS  ////////////////////////////

Sub Metals_Hit (idx)
        RandomSoundMetal
End Sub

Sub ShooterDiverter_collide(idx)
        RandomSoundMetal
End Sub

'/////////////////////////////  BOTTOM ARCH BALL GUIDE  ////////////////////////////
'/////////////////////////////  BOTTOM ARCH BALL GUIDE - SOFT BOUNCES  ////////////////////////////
Sub RandomSoundBottomArchBallGuide()
         dim finalspeed
          finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
         If finalspeed > 16 then
                PlaySoundAtLevelActiveBall ("Apron_Bounce_"& Int(Rnd*2)+1), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
        End if
        If finalspeed >= 6 AND finalspeed <= 16 then
                 Select Case Int(Rnd*2)+1
                        Case 1 : PlaySoundAtLevelActiveBall ("Apron_Bounce_1"), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
                        Case 2 : PlaySoundAtLevelActiveBall ("Apron_Bounce_Soft_1"), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
                End Select
         End If
        If finalspeed < 6 Then
                 Select Case Int(Rnd*2)+1
                        Case 1 : PlaySoundAtLevelActiveBall ("Apron_Bounce_Soft_1"), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
                        Case 2 : PlaySoundAtLevelActiveBall ("Apron_Medium_3"), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
                End Select
        End if
End Sub

'/////////////////////////////  BOTTOM ARCH BALL GUIDE - HARD HITS  ////////////////////////////
Sub RandomSoundBottomArchBallGuideHardHit()
        PlaySoundAtLevelActiveBall ("Apron_Hard_Hit_" & Int(Rnd*3)+1), BottomArchBallGuideSoundFactor * 0.25
End Sub

Sub Apron_Hit (idx)
        If Abs(cor.ballvelx(activeball.id) < 4) and cor.ballvely(activeball.id) > 7 then
                RandomSoundBottomArchBallGuideHardHit()
        Else
                RandomSoundBottomArchBallGuide
        End If
End Sub

'/////////////////////////////  FLIPPER BALL GUIDE  ////////////////////////////
Sub RandomSoundFlipperBallGuide()
         dim finalspeed
          finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
         If finalspeed > 16 then
                 Select Case Int(Rnd*2)+1
                        Case 1 : PlaySoundAtLevelActiveBall ("Apron_Hard_1"),  Vol(ActiveBall) * FlipperBallGuideSoundFactor
                        Case 2 : PlaySoundAtLevelActiveBall ("Apron_Hard_2"),  Vol(ActiveBall) * 0.8 * FlipperBallGuideSoundFactor
                End Select
        End if
        If finalspeed >= 6 AND finalspeed <= 16 then
                PlaySoundAtLevelActiveBall ("Apron_Medium_" & Int(Rnd*3)+1),  Vol(ActiveBall) * FlipperBallGuideSoundFactor
         End If
        If finalspeed < 6 Then
                PlaySoundAtLevelActiveBall ("Apron_Soft_" & Int(Rnd*7)+1),  Vol(ActiveBall) * FlipperBallGuideSoundFactor
        End If
End Sub

'/////////////////////////////  TARGET HIT SOUNDS  ////////////////////////////
Sub RandomSoundTargetHitStrong()
        PlaySoundAtLevelActiveBall SoundFX("Target_Hit_" & Int(Rnd*4)+5,DOFTargets), Vol(ActiveBall) * 0.45 * TargetSoundFactor
End Sub

Sub RandomSoundTargetHitWeak()
        PlaySoundAtLevelActiveBall SoundFX("Target_Hit_" & Int(Rnd*4)+1,DOFTargets), Vol(ActiveBall) * TargetSoundFactor
End Sub

Sub PlayTargetSound()
         dim finalspeed
          finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
         If finalspeed > 10 then
                 RandomSoundTargetHitStrong()
                RandomSoundBallBouncePlayfieldSoft Activeball
        Else
                 RandomSoundTargetHitWeak()
         End If
End Sub

Sub Targets_Hit (idx)
        PlayTargetSound
    TargetBouncer Activeball, 1
End Sub

'/////////////////////////////  BALL BOUNCE SOUNDS  ////////////////////////////
Sub RandomSoundBallBouncePlayfieldSoft(aBall)
        Select Case Int(Rnd*9)+1
                Case 1 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_1"), volz(aBall) * BallBouncePlayfieldSoftFactor, aBall
                Case 2 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_2"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.5, aBall
                Case 3 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_3"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.8, aBall
                Case 4 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_4"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.5, aBall
                Case 5 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_5"), volz(aBall) * BallBouncePlayfieldSoftFactor, aBall
                Case 6 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_1"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.2, aBall
                Case 7 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_2"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.2, aBall
                Case 8 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_5"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.2, aBall
                Case 9 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_7"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.3, aBall
        End Select
End Sub

Sub RandomSoundBallBouncePlayfieldHard(aBall)
        PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_" & Int(Rnd*7)+1), volz(aBall) * BallBouncePlayfieldHardFactor, aBall
End Sub

'/////////////////////////////  DELAYED DROP - TO PLAYFIELD - SOUND  ////////////////////////////
Sub RandomSoundDelayedBallDropOnPlayfield(aBall)
        Select Case Int(Rnd*5)+1
                Case 1 : PlaySoundAtLevelStatic ("Ball_Drop_Playfield_1_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
                Case 2 : PlaySoundAtLevelStatic ("Ball_Drop_Playfield_2_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
                Case 3 : PlaySoundAtLevelStatic ("Ball_Drop_Playfield_3_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
                Case 4 : PlaySoundAtLevelStatic ("Ball_Drop_Playfield_4_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
                Case 5 : PlaySoundAtLevelStatic ("Ball_Drop_Playfield_5_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
        End Select
End Sub

'/////////////////////////////  BALL GATES AND BRACKET GATES SOUNDS  ////////////////////////////

Sub SoundPlayfieldGate()
        PlaySoundAtLevelStatic ("Gate_FastTrigger_" & Int(Rnd*2)+1), GateSoundLevel, Activeball
End Sub

Sub SoundHeavyGate()
        PlaySoundAtLevelStatic ("Gate_2"), GateSoundLevel, Activeball
End Sub

Sub Gates_hit(idx)
        SoundHeavyGate
End Sub

Sub GatesWire_hit(idx)
        SoundPlayfieldGate
End Sub

'/////////////////////////////  LEFT LANE ENTRANCE - SOUNDS  ////////////////////////////

Sub RandomSoundLeftArch()
        PlaySoundAtLevelActiveBall ("Arch_L" & Int(Rnd*4)+1), Vol(ActiveBall) * ArchSoundFactor
End Sub

Sub RandomSoundRightArch()
        PlaySoundAtLevelActiveBall ("Arch_R" & Int(Rnd*4)+1), Vol(ActiveBall) * ArchSoundFactor
End Sub


Sub Arch1_hit()
        If Activeball.velx > 1 Then SoundPlayfieldGate
        StopSound "Arch_L1"
        StopSound "Arch_L2"
        StopSound "Arch_L3"
        StopSound "Arch_L4"
End Sub

Sub Arch1_unhit()
        If activeball.velx < -8 Then
                RandomSoundRightArch
        End If
End Sub

Sub Arch2_hit()
        If Activeball.velx < 1 Then SoundPlayfieldGate
        StopSound "Arch_R1"
        StopSound "Arch_R2"
        StopSound "Arch_R3"
        StopSound "Arch_R4"
End Sub

Sub Arch2_unhit()
        If activeball.velx > 10 Then
                RandomSoundLeftArch
        End If
End Sub

'/////////////////////////////  SAUCERS (KICKER HOLES)  ////////////////////////////

Sub SoundSaucerLock()
        PlaySoundAtLevelStatic ("Saucer_Enter_" & Int(Rnd*2)+1), SaucerLockSoundLevel, Activeball
End Sub

Sub SoundSaucerKick(scenario, saucer)
        Select Case scenario
                Case 0: PlaySoundAtLevelStatic SoundFX("Saucer_Empty", DOFContactors), SaucerKickSoundLevel, saucer
                Case 1: PlaySoundAtLevelStatic SoundFX("Saucer_Kick", DOFContactors), SaucerKickSoundLevel, saucer
        End Select
End Sub

'/////////////////////////////  BALL COLLISION SOUND  ////////////////////////////
Sub OnBallBallCollision(ball1, ball2, velocity)
        Dim snd
        Select Case Int(Rnd*7)+1
                Case 1 : snd = "Ball_Collide_1"
                Case 2 : snd = "Ball_Collide_2"
                Case 3 : snd = "Ball_Collide_3"
                Case 4 : snd = "Ball_Collide_4"
                Case 5 : snd = "Ball_Collide_5"
                Case 6 : snd = "Ball_Collide_6"
                Case 7 : snd = "Ball_Collide_7"
        End Select

        PlaySound (snd), 0, Csng(velocity) ^2 / 200 * BallWithBallCollisionSoundFactor * VolumeDial, AudioPan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub


'/////////////////////////////////////////////////////////////////
'                                        End Mechanical Sounds
'/////////////////////////////////////////////////////////////////



'******************************************************
'                BALL ROLLING AND DROP SOUNDS
'******************************************************

Const tnob = 10 ' total number of balls
ReDim rolling(tnob)
InitRolling

Dim DropCount
ReDim DropCount(tnob)

Sub InitRolling
        Dim i
        For i = 0 to tnob
                rolling(i) = False
        Next
End Sub

Sub RollingUpdate()
        Dim BOT, b
        BOT = GetBalls

        ' stop the sound of deleted balls
        For b = UBound(BOT) + 1 to tnob
                rolling(b) = False
                StopSound("BallRoll_" & b)
        Next

        ' exit the sub if no balls on the table
        If UBound(BOT) = -1 Then Exit Sub

        ' play the rolling sound for each ball

        For b = 0 to UBound(BOT)
                If BallVel(BOT(b)) > 1 AND BOT(b).z < 30 Then
                        rolling(b) = True
                        PlaySound ("BallRoll_" & b), -1, VolPlayfieldRoll(BOT(b)) * 1.1 * VolumeDial, AudioPan(BOT(b)), 0, PitchPlayfieldRoll(BOT(b)), 1, 0, AudioFade(BOT(b))

                Else
                        If rolling(b) = True Then
                                StopSound("BallRoll_" & b)
                                rolling(b) = False
                        End If
                End If

                '***Ball Drop Sounds***
                If BOT(b).VelZ < -1 and BOT(b).z < 55 and BOT(b).z > 27 Then 'height adjust for ball drop sounds
                        If DropCount(b) >= 5 Then
                                DropCount(b) = 0
                                If BOT(b).velz > -7 Then
                                        RandomSoundBallBouncePlayfieldSoft BOT(b)
                                Else
                                        RandomSoundBallBouncePlayfieldHard BOT(b)
                                End If
                        End If
                End If
                If DropCount(b) < 5 Then
                        DropCount(b) = DropCount(b) + 1
                End If
        Next
End Sub


'********************************************************************************
' Flex DMD routines made possible by scutters' tutorials and scripts.
'********************************************************************************
DIm FlexDMDFont
Dim FlexDMDFontActiveScore
Dim FlexDMDScene
Dim ExternalEnabled

Dim LastScoreUpdated
Dim FlexDMD   ' the flex dmd display
DIm FlexDMDDict   ' a dictionary / lookup to convert segment display hex/int codes to characters
Dim L1Chars, L2Chars, L3Chars, L4Chars
Dim Line1Change, Line2Change, Line3Change, Line4Change, Flexpath
Dim placemil, place100k, place10k, placek, place100, place10, place1
Dim fso,curdir
Dim Scorearray
Dim tikiframe
Sub FlexDMD_Init() 'default/startup values
  tikiframe=0
  Scorearray=Array("b","b","b","b","b","b","b")
  'arrays to hold characters to display converted from segment codes
  L1Chars = Array("0","0","0","0","b","0","0")
  L2Chars = Array("0","0","0","0","b","0","0")
  L3Chars = Array("0","0","0","0","b","0","0")
  L4Chars = Array("0","0","0","0","b","0","0")
  LastScoreUpdated = 0
  FlexDictionary_Init
  Set fso = CreateObject("Scripting.FileSystemObject")
  curDir = fso.GetAbsolutePathName(".")
  FlexPath = "VPX."

  Set FlexDMD = CreateObject("FlexDMD.FlexDMD")
  If Not FlexDMD is Nothing Then

    FlexDMD.GameName = cGameName
    FlexDMD.RenderMode = 2
    FlexDMD.Width = 128
    FlexDMD.Height = 32
    FlexDMD.Clear = True
    FlexDMD.Run = True
    FlexDMD.TableFile = Table1.Filename & ".vpx"
    Set FlexDMDScene = FlexDMD.NewGroup("Scene")


    With FlexDMDScene


      .AddActor FlexDMD.NewImage("Back", "VPX.background")
      .AddActor FlexDMD.NewImage("Movieframe",FlexPath & "glowsprites&region=0," & ((2) * 32) & ",128,32")
      .GetImage("Movieframe").Visible = False

      .AddActor FlexDMD.NewImage("Million",FlexPath & "0")
      .GetImage("Million").SetAlignedPosition  4,0,0
      .GetImage("Million").Visible = True

      .AddActor FlexDMD.Newimage("100K",FlexPath & "1")
      .Getimage("100K").SetAlignedPosition  22,0,0
      .Getimage("100K").Visible = True

      .AddActor FlexDMD.Newimage("Mcomma",FlexPath & "comma")
      .Getimage("Mcomma").SetAlignedPosition  20,28,0
      .Getimage("Mcomma").Visible = False

      .AddActor FlexDMD.Newimage("10K",FlexPath & "2")
      .Getimage("10K").SetAlignedPosition  40,0,0
      .Getimage("10K").Visible = True

      .AddActor FlexDMD.Newimage("1000",FlexPath & "3")
      .Getimage("1000").SetAlignedPosition  58,0,0
      .Getimage("1000").Visible = True

      .AddActor FlexDMD.Newimage("100",FlexPath & "4")
      .Getimage("100").SetAlignedPosition  76,0,0
      .Getimage("100").Visible = True

      .AddActor FlexDMD.Newimage("Kcomma",FlexPath & "comma")
      .Getimage("Kcomma").SetAlignedPosition  74,28,0
      .Getimage("Kcomma").Visible = False

      .AddActor FlexDMD.Newimage("10",FlexPath & "5")
      .Getimage("10").SetAlignedPosition  94,0,0
      .Getimage("10").Visible = True

      .AddActor FlexDMD.Newimage("1",FlexPath & "tiki&region=" & (tikiframe * 15) & ","&(0*32) &",15,32")
      .Getimage("1").SetAlignedPosition  112,0,0
      .Getimage("1").Visible = True

      .AddActor FlexDMD.NewImage("Title",FlexPath & "Title")
      .GetImage("Title").SetAlignedPosition  0,0,0
      .GetImage("Title").Visible = True

      .AddActor FlexDMD.NewImage("Pierre",FlexPath & "bumper1 1")
      .GetImage("Pierre").SetAlignedPosition  0,0,0
      .GetImage("Pierre").Visible = False

      .AddActor FlexDMD.NewImage("Jose",FlexPath & "bumper2 1")
      .GetImage("Jose").SetAlignedPosition  40,0,0
      .GetImage("Jose").Visible = False

      .AddActor FlexDMD.NewImage("Fritz",FlexPath & "bumper3 1")
      .GetImage("Fritz").SetAlignedPosition  88,0,0
      .GetImage("Fritz").Visible = False

      .addactor FlexDMD.Newimage("creditscroll", "VPX.credits&region=0,0,128,32")
      .GetImage("creditscroll").Visible = False

    End With

    FlexDMD.LockRenderThread

    FlexDMD.Stage.AddActor FlexDMDScene

    FlexDMD.Show = True
    FlexDMD.UnlockRenderThread

    Line1Change = False
    Line2Change = False
    Line3Change = False
    Line4Change = False
    changemovie
  End If

End Sub



Sub FlexDictionary_Init

  'add conversion of segment charcters codes to lookup table
  Set FlexDMDDict = CreateObject("Scripting.Dictionary")

  FlexDMDDict.Add 0, curdir & "b"
  FlexDMDDict.Add 63, curdir & "0"
  FlexDMDDict.Add 6, curdir & "1"
  FlexDMDDict.Add 91, curdir & "2"
  FlexDMDDict.Add 79, curdir & "3"
  FlexDMDDict.Add 102, curdir & "4"
  FlexDMDDict.Add 109, curdir & "5"
  FlexDMDDict.Add 125, curdir & "6"
  FlexDMDDict.Add 7, curdir & "7"
  FlexDMDDict.Add 127, curdir & "8"
  FlexDMDDict.Add 111, curdir & "9"

  FlexDMDDict.Add 191, curdir & "0"
  FlexDMDDict.Add 134, curdir & "1"
  FlexDMDDict.Add 219, curdir & "2"
  FlexDMDDict.Add 207, curdir & "3"
  FlexDMDDict.Add 230, curdir & "4"
  FlexDMDDict.Add 237, curdir & "5"
  FlexDMDDict.Add 253, curdir & "6"
  FlexDMDDict.Add 135, curdir & "7"
  FlexDMDDict.Add 255, curdir & "8"
  FlexDMDDict.Add 239, curdir & "9"

End Sub

'**************
' Update FlexDMD
'**************

Sub FlexDMDUpdate()

  if UseFlexDMD = 1 then
  If Not FlexDMD is Nothing Then FlexDMD.LockRenderThread
  If FlexDMD.Run = False Then FlexDMD.Run = True

    Dim i
      for i = 0 to 6
      If Line1Change Then
'     If L1Chars(i)<>"b" Then
Scorearray(i)= L1Chars(i)
      End If

      If Line2Change Then
'     If L2Chars(i)<>"b" Then
Scorearray(i)= L2Chars(i)
      End If

      If Line3Change Then
'     If L3Chars(i)<>"b" Then
Scorearray(i)= L3Chars(i)
      End If

      If Line4Change Then
'     If L4Chars(i)<>"b" Then
Scorearray(i)= L4Chars(i)
      End If
      Next

  If credittimer.Enabled = True Then
    for i = 0 to 6
      Scorearray(i)= L1Chars(i)
    Next
  End If



  With FlexDMD.Stage
    If Scorearray(0)="0" Then
    Scorearray(0)="b"
      If Scorearray(1) = "0" then
      Scorearray(1) = "b"
        If Scorearray(2) = "0"  then
        Scorearray(2) = "b"
          If Scorearray(3) = "0"  then
          Scorearray(3) = "b"
            If Scorearray(4) = "0" then
            Scorearray(4) = "b"
              'If Scorearray(5) = "0" then Scorearray(5) = "b"
            End If
          End If
        End If
      End If
    End If
' If creditson = true and Scorearray (6) <> "b" then exit Sub
    .GetImage("Title").Visible = False
    .GetImage("Million").Bitmap = FlexDMD.NewImage("Million",FlexPath & Scorearray(0) &"").Bitmap
    .GetImage("100K").Bitmap = FlexDMD.NewImage("Million",FlexPath & Scorearray(1) &"").Bitmap
    .GetImage("10K").Bitmap = FlexDMD.NewImage("Million",FlexPath & Scorearray(2) &"").Bitmap
    .GetImage("1000").Bitmap = FlexDMD.NewImage("Million",FlexPath & Scorearray(3) &"").Bitmap
    .GetImage("100").Bitmap = FlexDMD.NewImage("Million",FlexPath & Scorearray(4) &"").Bitmap
    .GetImage("10").Bitmap = FlexDMD.NewImage("Million",FlexPath & Scorearray(5) &"").Bitmap
    If (moviemode <= 5 or  moviemode= spriterows) Then
      If Scorearray(3) <> "b" and creditson = false Then
      .Getimage("Kcomma").Visible = True
      Else
      .Getimage("Kcomma").Visible = False
      End If
      If Scorearray(0) <>"b" Then
      .Getimage("Mcomma").Visible = True
      Else
      .Getimage("Mcomma").Visible = False
      End If
      If Scorearray(5) <>"b" and creditson = false Then
      .Getimage("1").Visible = True
      Else
      .Getimage("1").Visible = False
      End If
    End If

  End With

  If Not FlexDMD is Nothing Then FlexDMD.UnlockRenderThread

  Line1Change = False
  Line2Change = False
  Line3Change = False
  Line4Change = False
    End If
End Sub


Sub UpdateFlexChar(id, value)
  'map segment code to character in LnChars arrays
  Dim chr
  if id < 28 and FlexDMDDict.Exists (value) then

    chr = FlexDMDDict.Item (value)



    if id < 7 then
      L1Chars(id) = chr
      Line1Change = True
    elseif id < 14 then
      L2Chars(id - 7) = chr
      Line2Change = True
    elseif id < 21 then
      L3Chars(id - 14) = chr
      Line3Change = True
    elseif id < 28 then
      L4Chars(id - 21) = chr
      Line4Change = True
    end if

  end if

End Sub

Dim moviemode, framenumber
Dim bump1frame, bump2frame, bump3frame
Sub movemovie
  if UseFlexDMD = 1 then
  If Not FlexDMD is Nothing Then FlexDMD.LockRenderThread
  If FlexDMD.Run = False Then FlexDMD.Run = True
  if moviemode="" then moviemode=spriterows
  if framenumber="" then framenumber=0

  With FlexDMD.Stage

    If moviemode=spriterows then
    .GetImage("Movieframe").Visible = False
    Elseif moviemode=birdmode Then
    .GetImage("Movieframe").Visible = False
    bump1frame=bump1frame+1
    bump2frame=bump2frame+1
    bump3frame=bump3frame+1
      .GetImage("Pierre").Visible = True
      .GetImage("Jose").Visible = True
      .GetImage("Fritz").Visible = True

      If bump1frame = 1 or bump1frame=3 or bump1frame=5 Then
      .GetImage("Pierre").Bitmap = FlexDMD.NewImage("Pierre",FlexPath & "bumper1 2").bitmap
      Elseif bump1frame = 2 or bump1frame=4 Then
      .GetImage("Pierre").Bitmap = FlexDMD.NewImage("Pierre",FlexPath & "bumper1 3").bitmap
      Else
      .GetImage("Pierre").Bitmap = FlexDMD.NewImage("Pierre",FlexPath & "bumper1 1").bitmap
      End If

      If bump2frame = 1 or bump2frame=3 or bump2frame=5 Then
      .GetImage("Jose").Bitmap = FlexDMD.NewImage("Jose",FlexPath & "bumper2 2").bitmap
      Elseif bump2frame = 2 or bump2frame=4 Then
      .GetImage("Jose").Bitmap = FlexDMD.NewImage("Jose",FlexPath & "bumper2 3").bitmap
      Else
      .GetImage("Jose").Bitmap = FlexDMD.NewImage("Jose",FlexPath & "bumper2 1").bitmap
      End If

      If bump3frame = 1 or bump3frame=3 or bump3frame=5 Then
      .GetImage("Fritz").Bitmap = FlexDMD.NewImage("Fritz",FlexPath & "Bumper3 2").bitmap
      Elseif bump3frame = 2 or bump3frame=4 Then
      .GetImage("Fritz").Bitmap = FlexDMD.NewImage("Fritz",FlexPath & "Bumper3 3").bitmap
      Else
      .GetImage("Fritz").Bitmap = FlexDMD.NewImage("Fritz",FlexPath & "Bumper3 1").bitmap
      End If


    Else
    .GetImage("Movieframe").Bitmap = FlexDMD.NewImage("Movieframe",FlexPath & "glowsprites&region=" & (framenumber*128)& ","&(moviemode * 32) & ",128,32").Bitmap
    .GetImage("Movieframe").Visible = True
    End If
  End With
  framenumber=framenumber+1
  if framenumber=10 then
    framenumber=0
    Select Case moviemode
    Case 0,1,2,3,4,5, spriterows : changemovie
    Case 6: moviemode=7
    Case 7: moviemode=8
    Case 8: digitson: changemovie
    Case 9: moviemode=10
    Case 10: digitson: changemovie
    Case 11:If volcanocounter = "" Then
        volcanocounter = 1
        Else volcanocounter = ""
        changemovie
        digitson
        End If
    Case 12: digitson: changemovie
    Case 13: digitson: changemovie
    Case 14: digitson: changemovie
    Case 15: moviemode=18 :splash
    Case 16: moviemode = 18:splash
    Case 17: moviemode = 18:splash
    Case 18
    if girl = true Then
      girl = False
      if cow=True then
      moviemode = 17: moo
      Else
      moviemode = 15
      end If
    Elseif cow=true then
      moviemode = 15
      cow= False
      Else
      digitson
      changemovie
      End If
    Case 19: digitson: hidebirds :changemovie
    Case birdmode: digitson: hidebirds :changemovie
    End Select
  end if
  If Not FlexDMD is Nothing Then FlexDMD.UnlockRenderThread
  End If
End Sub
Dim cow, girl
cow = False
girl=false
Const spriterows = 21
Const birdmode = 20 'must be one less than spriterows
Dim volcanocounter

sub changemovie
  hidebirds
  moviemode= Int(rnd*20)
  if moviemode>5 then moviemode=spriterows
end sub

Sub digitsoff
  if UseFlexDMD = 1 then
  If Not FlexDMD is Nothing Then FlexDMD.LockRenderThread
  If FlexDMD.Run = False Then FlexDMD.Run = True
  With FlexDMD.Stage
  .GetImage("Million").Visible = False
  .GetImage("100K").Visible = False
  .GetImage("10K").Visible = False
  .GetImage("1000").Visible = False
  .GetImage("100").Visible = False
  .GetImage("10").Visible = False
  .GetImage("1").Visible = False
  .Getimage("Mcomma").Visible = False
  .Getimage("Kcomma").Visible = False
  End With
  If Not FlexDMD is Nothing Then FlexDMD.UnlockRenderThread
  End If
End Sub


Sub digitson
  if UseFlexDMD = 1 then
  If Not FlexDMD is Nothing Then FlexDMD.LockRenderThread
  If FlexDMD.Run = False Then FlexDMD.Run = True
  With FlexDMD.Stage
  If Scorearray(0)<>"0" Then.GetImage("Million").Visible = True
  .GetImage("100K").Visible = True
  .GetImage("10K").Visible = True
  .GetImage("1000").Visible = True
  .GetImage("100").Visible = True
  .GetImage("10").Visible = True
  .GetImage("1").Visible = True
  .GetImage("creditscroll").Visible = False
  End With
  If Not FlexDMD is Nothing Then FlexDMD.UnlockRenderThread
End If
End Sub

Sub hidebirds
  if UseFlexDMD = 1 then
  If Not FlexDMD is Nothing Then FlexDMD.LockRenderThread
  If FlexDMD.Run = False Then FlexDMD.Run = True
  With FlexDMD.Stage
      .GetImage("Pierre").Visible = False
      .GetImage("Jose").Visible = False
      .GetImage("Fritz").Visible = False
  End With
  If Not FlexDMD is Nothing Then FlexDMD.UnlockRenderThread
  End If
End Sub

Sub tikiupdate
  if UseFlexDMD = 1 then
  tikiframe=tikiframe+1: if tikiframe >9 then tikiframe=0
  If Not FlexDMD is Nothing Then FlexDMD.LockRenderThread
  If FlexDMD.Run = False Then FlexDMD.Run = True
  With FlexDMD.Stage
  .GetImage("1").Bitmap =FlexDMD.Newimage("1",FlexPath & "tiki&region=" & (tikiframe * 15) & ","&(0*32) &",15,32").Bitmap
' .GetImage("Movieframe").Bitmap = FlexDMD.NewImage("Movieframe",FlexPath & "glowsprites&region=" & (framenumber*128)& ","&(moviemode * 32) & ",128,32").Bitmap

  End With
  If Not FlexDMD is Nothing Then FlexDMD.UnlockRenderThread
End If
End Sub

Sub splash

  Dim xx
  xx =  INT(3 * RND(1) )
  If UseFlexDMD = 1 then PlaySound ("splash"& xx)
End Sub
Sub moo
  Dim xx
  xx =  INT(4 * RND(1) )
  If UseFlexDMD = 1 then  PlaySound ("cow"& xx)
End Sub
Sub giggle
' Dim xx
' xx =  INT(4 * RND(1) )
  If UseFlexDMD then  PlaySound"girl0",0,0.3
End Sub


Sub scrollcredits
  If UseFlexDMD = 1 then
  If Not FlexDMD is Nothing Then FlexDMD.LockRenderThread
  If FlexDMD.Run = False Then FlexDMD.Run = True
  With FlexDMD.Stage

  digitsoff
  .GetImage("creditscroll").Bitmap =FlexDMD.Newimage("creditscroll", "VPX.credits&region=0," & (creditframe) & ",128,32").Bitmap
  .GetImage("creditscroll").Visible = True

  End With
  If Not FlexDMD is Nothing Then FlexDMD.UnlockRenderThread
  End If
End Sub

dim creditframe
creditframe= -200
dim creditson
creditson=False
Sub credittimer_Timer()
  creditframe=creditframe+1
  If creditframe >-1 Then
  scrollcredits
  creditson = True
    Else creditson = false
  End If
  if creditframe = 1367 Then  creditframe=-800

End Sub

'******************************
' Setup VR Backglass & Digits
'******************************

Dim xoff,yoff1, yoff2, yoff3, yoff4, yoff5,zoff,xrot,zscale, xcen,ycen

Sub setup_backglass()

  xoff = -20
  yoff1 = 153 ' this is where you adjust the forward/backward position for player 1 score
  yoff2 = 153 ' this is where you adjust the forward/backward position for player 2 score
  yoff3 = 153 ' this is where you adjust the forward/backward position for player 3 score
  yoff4 = 153 ' this is where you adjust the forward/backward position for player 4 score
  yoff5 = 153 ' this is where you adjust the forward/backward position for credits and ball in play
  zoff = 699
  xrot = -90

  center_digits()

End Sub


Sub center_digits()
  Dim ix, xx, yy, yfact, xfact, xobj

  zscale = 0.0000001

  xcen = (130 /2) - (92 / 2)
  ycen = (780 /2 ) + (203 /2)

  for ix = 0 to 6
    For Each xobj In vDigits(ix)

      xx = xobj.x

      xobj.x = (xoff - xcen) + xx
      yy = xobj.y ' get the yoffset before it is changed
      xobj.y = yoff1

      If (yy < 0.) then
        yy = yy * -1
      end if

      xobj.height = (zoff - ycen) + yy - (yy * (zscale))
      xobj.rotx = xrot
    Next
  Next

  for ix = 7 to 13
    For Each xobj In vDigits(ix)

      xx = xobj.x

      xobj.x = (xoff - xcen) + xx
      yy = xobj.y ' get the yoffset before it is changed
      xobj.y = yoff2

      If (yy < 0.) then
        yy = yy * -1
      end if

      xobj.height = (zoff - ycen) + yy - (yy * (zscale))
      xobj.rotx = xrot
    Next
  Next

  for ix = 14 to 20
    For Each xobj In vDigits(ix)

      xx = xobj.x

      xobj.x = (xoff - xcen) + xx
      yy = xobj.y ' get the yoffset before it is changed
      xobj.y = yoff3

      If (yy < 0.) then
        yy = yy * -1
      end if

      xobj.height = (zoff - ycen) + yy - (yy * (zscale))
      xobj.rotx = xrot
    Next
  Next

  for ix = 21 to 27
    For Each xobj In vDigits(ix)

      xx = xobj.x

      xobj.x = (xoff - xcen) + xx
      yy = xobj.y ' get the yoffset before it is changed
      xobj.y = yoff4

      If (yy < 0.) then
        yy = yy * -1
      end if

      xobj.height = (zoff - ycen) + yy - (yy * (zscale))
      xobj.rotx = xrot
    Next
  Next

  for ix = 28 to 31
    For Each xobj In vDigits(ix)

      xx = xobj.x

      xobj.x = (xoff - xcen) + xx
      yy = xobj.y ' get the yoffset before it is changed
      xobj.y = yoff5

      If (yy < 0.) then
        yy = yy * -1
      end if

      xobj.height = (zoff - ycen) + yy - (yy * (zscale))
      xobj.rotx = xrot
    Next
  Next

End Sub

Dim vDigits(32)
vDigits(0) = Array(LED1x0,LED1x1,LED1x2,LED1x3,LED1x4,LED1x5,LED1x6)
vDigits(1) = Array(LED2x0,LED2x1,LED2x2,LED2x3,LED2x4,LED2x5,LED2x6)
vDigits(2) = Array(LED3x0,LED3x1,LED3x2,LED3x3,LED3x4,LED3x5,LED3x6)
vDigits(3) = Array(LED4x0,LED4x1,LED4x2,LED4x3,LED4x4,LED4x5,LED4x6)
vDigits(4) = Array(LED5x0,LED5x1,LED5x2,LED5x3,LED5x4,LED5x5,LED5x6)
vDigits(5) = Array(LED6x0,LED6x1,LED6x2,LED6x3,LED6x4,LED6x5,LED6x6)
vDigits(6) = Array(LED7x0,LED7x1,LED7x2,LED7x3,LED7x4,LED7x5,LED7x6)

vDigits(7) = Array(LED8x0,LED8x1,LED8x2,LED8x3,LED8x4,LED8x5,LED8x6)
vDigits(8) = Array(LED9x0,LED9x1,LED9x2,LED9x3,LED9x4,LED9x5,LED9x6)
vDigits(9) = Array(LED10x0,LED10x1,LED10x2,LED10x3,LED10x4,LED10x5,LED10x6)
vDigits(10) = Array(LED11x0,LED11x1,LED11x2,LED11x3,LED11x4,LED11x5,LED11x6)
vDigits(11) = Array(LED12x0,LED12x1,LED12x2,LED12x3,LED12x4,LED12x5,LED12x6)
vDigits(12) = Array(LED13x0,LED13x1,LED13x2,LED13x3,LED13x4,LED13x5,LED13x6)
vDigits(13) = Array(LED14x0,LED14x1,LED14x2,LED14x3,LED14x4,LED14x5,LED14x6)

vDigits(14) = Array(LED1x000,LED1x001,LED1x002,LED1x003,LED1x004,LED1x005,LED1x006)
vDigits(15) = Array(LED1x100,LED1x101,LED1x102,LED1x103,LED1x104,LED1x105,LED1x106)
vDigits(16) = Array(LED1x200,LED1x201,LED1x202,LED1x203,LED1x204,LED1x205,LED1x206)
vDigits(17) = Array(LED1x300,LED1x301,LED1x302,LED1x303,LED1x304,LED1x305,LED1x306)
vDigits(18) = Array(LED1x400,LED1x401,LED1x402,LED1x403,LED1x404,LED1x405,LED1x406)
vDigits(19) = Array(LED1x500,LED1x501,LED1x502,LED1x503,LED1x504,LED1x505,LED1x506)
vDigits(20) = Array(LED1x600,LED1x601,LED1x602,LED1x603,LED1x604,LED1x605,LED1x606)

vDigits(21) = Array(LED2x000,LED2x001,LED2x002,LED2x003,LED2x004,LED2x005,LED2x006)
vDigits(22) = Array(LED2x100,LED2x101,LED2x102,LED2x103,LED2x104,LED2x105,LED2x106)
vDigits(23) = Array(LED2x200,LED2x201,LED2x202,LED2x203,LED2x204,LED2x205,LED2x206)
vDigits(24) = Array(LED2x300,LED2x301,LED2x302,LED2x303,LED2x304,LED2x305,LED2x306)
vDigits(25) = Array(LED2x400,LED2x401,LED2x402,LED2x403,LED2x404,LED2x405,LED2x406)
vDigits(26) = Array(LED2x500,LED2x501,LED2x502,LED2x503,LED2x504,LED2x505,LED2x506)
vDigits(27) = Array(LED2x600,LED2x601,LED2x602,LED2x603,LED2x604,LED2x605,LED2x606)

vDigits(28) = Array(LEDax300,LEDax301,LEDax302,LEDax303,LEDax304,LEDax305,LEDax306)
vDigits(29) = Array(LEDbx400,LEDbx401,LEDbx402,LEDbx403,LEDbx404,LEDbx405,LEDbx406)
vDigits(30) = Array(LEDcx500,LEDcx501,LEDcx502,LEDcx503,LEDcx504,LEDcx505,LEDcx506)
vDigits(31) = Array(LEDdx600,LEDdx601,LEDdx602,LEDdx603,LEDdx604,LEDdx605,LEDdx606)

dim DisplayColor
DisplayColor =  RGB(152,255,0)

Sub VRDigitDisplay
    Dim ChgLED, ii, jj, num, chg, stat, obj, b, x
    ChgLED=Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
    If Not IsEmpty(ChgLED)Then
       For ii=0 To UBound(chgLED)
          num=chgLED(ii, 0) : chg=chgLED(ii, 1) : stat=chgLED(ii, 2)
      If UseFlexDMD = 1 then UpdateFlexChar num, stat
              For Each obj In vDigits(num)
'                   If chg And 1 Then obj.visible=stat And 1    'if you use the object color for off; turn the display object visible to not visible on the playfield, and uncomment this line out.
           If chg And 1 Then FadeDisplay obj, stat And 1
                   chg=chg\2 : stat=stat\2
              Next
        Next
    If UseFlexDMD = 1 then FlexDMDUpdate
    End If
 End Sub

Sub FadeDisplay(object, onoff)
  If OnOff = 1 Then
    object.color = DisplayColor
    Object.Opacity = 12
  Else
    Object.Color = RGB(1,1,1)
    Object.Opacity = 6
  End If
End Sub

Sub InitDigits()
  dim tmp, x, obj
  For x = 0 to uBound(vDigits)
    If IsArray(vDigits(x) ) then
      For each obj in vDigits(x)
        obj.height = obj.height + 18
        FadeDisplay obj, 0
      next
    End If
  Next
End Sub

InitDigits

' ******************************************************************************************
'      LAMP CALLBACK for the 6 backglass flasher lamps (not the solenoid conrolled ones)
' ******************************************************************************************

Sub UpdateMultipleLamps()
  If Controller.Lamp(63) = 0 Then: FlBGL63.visible=0: else: FlBGL63.visible=1 'Match
  If Controller.Lamp(61) = 0 Then: FlBGL61.visible=0: else: FlBGL61.visible=1 'Tilt
  If Controller.Lamp(45) = 0 Then: FlBGL45.visible=0: else: FlBGL45.visible=1 'Game Over
  If Controller.Lamp(43) = 0 Then: FlBGL43.visible=0: else: FlBGL43.visible=1 'Same Player Shoots Again
  If Controller.Lamp(13) = 0 Then: FlBGL13.visible=0: else: FlBGL13.visible=1 'High Score
End Sub

'*****************************************************************************************************
' VR Plunger Animation
'*****************************************************************************************************
Sub TimerPlunger_Timer
  If Pincab_Plunger.Y < 2349 then
    Pincab_Plunger.Y = Pincab_Plunger.Y + 5
  End If
End Sub

Sub TimerPlunger2_Timer
  Pincab_Plunger.Y = 2241 + (5 * Plunger.Position) - 20
End Sub

'vr bg animations

dim vran1st: vran1st = 1

Sub vrBGAn1_timer()

  Select Case vran1st

    Case 1: VRBG_Blaze1.visible = 1:VRBG_Blaze2.visible = 0:VRBG_Blaze3.visible = 0:VRBG_Volcano.visible = 1

    Case 2: VRBG_Blaze1.visible = 1:VRBG_Blaze2.visible = 0:VRBG_Blaze3.visible = 0:VRBG_Flame1.visible = 1:VRBG_Flame2.visible = 0:VRBG_Flame3.visible = 1:VRBG_Volcano.visible = 0

    Case 4: VRBG_Blaze1.visible = 1:VRBG_Blaze2.visible = 0:VRBG_Blaze3.visible = 0:VRBG_Flame1.visible = 0:VRBG_Flame2.visible = 1:VRBG_Flame3.visible = 1:VRBG_Volcano.visible = 0

    Case 5: VRBG_Blaze1.visible = 0:VRBG_Blaze2.visible = 1:VRBG_Blaze3.visible = 0:VRBG_Volcano.visible = 1

    Case 6: VRBG_Blaze1.visible = 1:VRBG_Blaze2.visible = 0:VRBG_Blaze3.visible = 0:VRBG_Flame1.visible = 1:VRBG_Flame2.visible = 0:VRBG_Flame3.visible = 0:VRBG_Volcano.visible = 0

    Case 8: VRBG_Blaze1.visible = 1:VRBG_Blaze2.visible = 0:VRBG_Blaze3.visible = 0:VRBG_Flame1.visible = 0:VRBG_Flame2.visible = 1:VRBG_Flame3.visible = 0:VRBG_Volcano.visible = 0

    Case 10: VRBG_Blaze1.visible = 1:VRBG_Blaze2.visible = 0:VRBG_Blaze3.visible = 1:VRBG_Flame1.visible = 1:VRBG_Flame2.visible = 0:VRBG_Flame3.visible = 1:VRBG_Volcano.visible = 1

    Case 12: VRBG_Blaze1.visible = 1:VRBG_Blaze2.visible = 0:VRBG_Blaze3.visible = 0:VRBG_Flame1.visible = 0:VRBG_Flame2.visible = 1:VRBG_Flame3.visible = 1:VRBG_Volcano.visible = 0

    Case 14: VRBG_Blaze1.visible = 1:VRBG_Blaze2.visible = 0:VRBG_Blaze3.visible = 0:VRBG_Flame1.visible = 1:VRBG_Flame2.visible = 0:VRBG_Flame3.visible = 0:VRBG_Volcano.visible = 0

    Case 15: VRBG_Blaze1.visible = 0:VRBG_Blaze2.visible = 1:VRBG_Blaze3.visible = 1:VRBG_Volcano.visible = 1

    Case 16: VRBG_Blaze1.visible = 1:VRBG_Blaze2.visible = 0:VRBG_Blaze3.visible = 0:VRBG_Flame1.visible = 0:VRBG_Flame2.visible = 1:VRBG_Flame3.visible = 0:VRBG_Volcano.visible = 0

    Case 18: VRBG_Blaze1.visible = 1:VRBG_Blaze2.visible = 0:VRBG_Blaze3.visible = 0:VRBG_Flame1.visible = 1:VRBG_Flame2.visible = 0:VRBG_Flame3.visible = 1:VRBG_Volcano.visible = 0

    Case 20: VRBG_Blaze1.visible = 1:VRBG_Blaze2.visible = 0:VRBG_Blaze3.visible = 1:VRBG_Flame1.visible = 0:VRBG_Flame2.visible = 1:VRBG_Flame3.visible = 1:VRBG_Volcano.visible = 2

    Case 22: VRBG_Blaze1.visible = 1:VRBG_Blaze2.visible = 0:VRBG_Blaze3.visible = 0:VRBG_Flame1.visible = 1:VRBG_Flame2.visible = 0:VRBG_Flame3.visible = 0:VRBG_Volcano.visible = 0

    Case 24: VRBG_Blaze1.visible = 1:VRBG_Blaze2.visible = 0:VRBG_Blaze3.visible = 0:VRBG_Flame1.visible = 0:VRBG_Flame2.visible = 1:VRBG_Flame3.visible = 0:VRBG_Volcano.visible = 0

    Case 25: VRBG_Blaze1.visible = 0:VRBG_Blaze2.visible = 1:VRBG_Blaze3.visible = 0:VRBG_Volcano.visible = 1

    Case 26: VRBG_Blaze1.visible = 1:VRBG_Blaze2.visible = 0:VRBG_Blaze3.visible = 0:VRBG_Flame1.visible = 1:VRBG_Flame2.visible = 0:VRBG_Flame3.visible = 1:VRBG_Volcano.visible = 0

    Case 28: VRBG_Blaze1.visible = 1:VRBG_Blaze2.visible = 0:VRBG_Blaze3.visible = 0:VRBG_Flame1.visible = 0:VRBG_Flame2.visible = 1:VRBG_Flame3.visible = 1:VRBG_Volcano.visible = 0

    Case 30: VRBG_Blaze1.visible = 1:VRBG_Blaze2.visible = 0:VRBG_Blaze3.visible = 0:VRBG_Flame1.visible = 1:VRBG_Flame2.visible = 0:VRBG_Flame3.visible = 0:VRBG_Volcano.visible = 1


  End Select

  vran1st = vran1st + 1

  if vran1st = 31 Then

    vran1st = 1


  end If
end sub

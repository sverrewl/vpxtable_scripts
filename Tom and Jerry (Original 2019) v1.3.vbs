'############################################################################################
'#######                                                                             ########
'#######          Tom & Jerry                                                        ########
'#######          (Original Mod of Hollywood Heat by watacaractr)                    ########
'#######                                                                             ########
'############################################################################################

' Version 1.0 - watacaractr - 2018
' - Original version which includes all base graphics, primitives, meshes, logic, scripting, etc.  A beautiful rendition of a ROM mod.  Thanks watacaractr!

' Version 1.2 - Gedankekojote97 - 2022
' - Added nFozzy physics, Fleep sounds and LUT selector.

' Version 1.3 - TastyWasps - 2023
' - Added a hybrid VR Room with cabinet artwork for VR.  Desktop background upgraded to include Tom & Jerry logo.

Option Explicit
Randomize

Const cGameName = "tomjerry"

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "Can't open controller.vbs"
On Error Goto 0

LoadVPM "01210000", "sys80.VBS", 3.1

'**********************************************************
'********       OPTIONS     *******************************
'**********************************************************

Dim BallShadows: Ballshadows=1          '******************set to 1 to turn on Ball shadows
Dim FlipperShadows: FlipperShadows=1  '***********set to 1 to turn on Flipper shadows
Dim ROMSounds: ROMSounds=0        '**********set to 0 for no rom sounds, 1 to play rom sounds.. mostly used for testing

'----- Desktop, VR, or Full Cabinet Options -----
Dim VR_Room, VR_Obj

If RenderingMode = 2 Then
  VR_Room = 1
Else
  VR_Room = 0
End If

If VR_Room = 1 Then
  For Each VR_Obj in VR : VR_Obj.Visible = 1 : Next
  SideRailLeft.Visible = 0
  SideRailRight.Visible = 0
Else
  For Each VR_Obj in VR : VR_Obj.Visible = 0 : Next
End If

'************************************************
'************************************************
'************************************************
'************************************************
'************************************************
Const UseSolenoids = True
Const UseLamps = True
Const UseSync = False
Const UseGI = False

' Standard Sounds
Const SSolenoidOn = "fx_solenoid"
Const SSolenoidOff = "fx_solenoidoff"
Const SCoin = "fx_coin"

Dim bsTrough, bslLock, bsrLock, dtRight, dtLeft, FastFlips, objekt, xx
Dim bMusicOn, startgamesound, multiballs

Sub TomandJerry_Init

' Thalamus : Was missing 'vpminit me'
  vpminit me

     With Controller
         .GameName = cGameName
         If Err Then MsgBox "Can't start Game " & cGameName & vbNewLine & Err.Description:Exit Sub
         .SplashInfoLine = "Tom & Jerry 1.2"&chr(13)&"Williams 2018"
         .HandleKeyboard = 0
         .ShowTitle = 0
         .ShowDMDOnly = 1
         .ShowFrame = 0
         .HandleMechanics = False
     .Games(cGameName).Settings.Value("sound") = ROMSounds
     .Hidden = 0
         On Error Resume Next
         .Run GetPlayerHWnd
         If Err Then MsgBox Err.Description
         On Error Goto 0
     End With

  If B2SOn then
    for each objekt in backdropstuff : objekt.visible = 0 : next
  End If

  Intensity 'sets GI brightness depending on day/night slider settings

  multiballs=1
  ALlightsTimer.uservalue=1

  CaptiveKick.createball
  CaptiveKick.kick 0,0

    PinMAMETimer.Interval = PinMAMEInterval
    PinMAMETimer.Enabled = 1

'************************Trough

  sw76.CreateBall
  sw20.CreateBall
  sw10.CreateBall
  Controller.Switch(76) = 1
  Controller.Switch(20) = 1
  Controller.Switch(10) = 1

'Kickers

    Set bslLock=New cvpmBallStack
    with bslLock
        .InitSaucer sw46,46,-170,10
        .InitExitSnd Soundfx("fx_ballrel",DOFContactors), Soundfx("HoleKick",DOFContactors)
    end with

    Set bsrLock=New cvpmBallStack
    with bsrLock
       .InitSaucer sw56,56,120,11
       .InitExitSnd Soundfx("fx_ballrel",DOFContactors), Soundfx("HoleKick",DOFContactors)
       .KickForceVar = 3
       .KickAngleVar = 3
    end with

' Nudging
  vpmNudge.TiltSwitch = 57
    vpmNudge.Sensitivity = 1
    vpmNudge.TiltObj = Array(leftslingshot, rightslingshot, TopSlingShot, URightSlingShot, bumper1)

    Set dtLeft=New cvpmDropTarget
        dtLeft.InitDrop Array(sw40,sw50,sw60),Array(40,50,60)
        dtLeft.InitSnd SoundFX("",DOFDropTargets),SoundFX("DTReset",DOFDropTargets)

    Set dtRight=New cvpmDropTarget
        dtRight.InitDrop Array(sw41,sw51,sw61),Array(41,51,61)
        dtRight.InitSnd SoundFX("",DOFDropTargets),SoundFX("DTReset",DOFDropTargets)

  if ballshadows=1 then
        BallShadowUpdate.enabled=1
      else
        BallShadowUpdate.enabled=0
    end if

    if flippershadows=1 then
        FlipperLSh.visible=1
        FlipperRSh.visible=1
        MiniFlipperLSh.visible=1
        MiniFlipperRSh.visible=1
      else
        FlipperLSh.visible=0
        FlipperRSh.visible=0
        MiniFlipperLSh.visible=0
        MiniFlipperRSh.visible=0
    end if

  vpmMapLights CPULights

    Set FastFlips = new cFastFlips
    with FastFlips
        .CallBackL = "SolLflipper"  'Point these to flipper subs
        .CallBackR = "SolRflipper"  '...
    '   .CallBackUL = "SolULflipper"'...(upper flippers, if needed)
    '   .CallBackUR = "SolURflipper"'...
        .TiltObjects = True 'Optional, if True calls vpmnudge.solgameon automatically. IF YOU GET A LINE 1 ERROR, DISABLE THIS! (or setup vpmNudge.TiltObj!)
    '   .DebugOn = False        'Debug, always-on flippers. Call FastFlips.DebugOn True or False in debugger to enable/disable.
    end with

  startgamesound=1

  PlaySound "T&Jambient", -1, .05     'adjust the volume using the second parameter.

End Sub

'***************
' Random Music
'***************

Dim song, numsongs, currentsong
numsongs=5    'Set to number of songs, name songs T&JbackgroundMusic0, T&JbackgroundMusic1, T&JbackgroundMusic2.. etc

Sub PlaySong 'random
    If bMusicOn Then
    currentsong = INT((numsongs-1) * RND(1) )
        song = "bgout_T&JbackgroundMusic" & currentsong & ".mp3"
        PlayMusic song
    End If
End Sub

Sub TomandJerry_MusicDone
    If bMusicOn Then
    if multiballs>1 then
      PlayMusic "bgout_T&Jmultiballmadness.mp3"     'if still in multiball repeat multiball music
      else
      PlaySong              'if not in multiball, play random song.
    end if
    End If
End Sub


'************************************************
' Solenoids
'************************************************
SolCallback(1) =    "SolLLock"
SolCallback(2) =    "SolRLock"
SolCallback(4) =    "dtDrop2"
SolCallback(5) =    "SolLeftTargetReset"
SolCallback(6) =    "SolRightTargetReset"
SolCallback(7) =     "dtDrop3"
SolCallback(8) =    "solknocker"
SolCallback(9) =    "solballrelease"
SolCallback(10) = "FastFlips.TiltSol"

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
     If Enabled Then
        PlaySound SoundFX("fx_Flipperup",DOFFlippers):LeftFlipper.RotateToEnd: LeftFlipper1.RotateToEnd
    controller.Switch(6)=1
     Else
        PlaySound SoundFX("fx_Flipperdown",DOFFlippers):LeftFlipper.RotateToStart: LeftFlipper1.RotateToStart
    controller.Switch(6)=0
     End If
  End Sub

' paired drop target dropping in "In-Sync" mode

 sub dtDrop2(enabled):dtLeft.hit 2:dtRight.hit 2:l7.state=0:end sub
 sub dtDrop3(enabled):dtLeft.hit 3:dtRight.hit 3:l8.state=0:end sub

Sub SolRFlipper(Enabled)
     If Enabled Then
         PlaySound SoundFX("fx_Flipperup",DOFFlippers):RightFlipper.RotateToEnd:RightFlipper1.RotateToEnd
    controller.Switch(16)=1
    controller.Switch(71)=1
     Else
         PlaySound SoundFX("fx_Flipperdown",DOFFlippers):RightFlipper.RotateToStart:RightFlipper1.RotateToStart
    controller.Switch(16)=0
    controller.Switch(71)=0
     End If
End Sub

Dim GILevel, DayNight

Sub Intensity
  If DayNight <= 20 Then
      GILevel = .5
  ElseIf DayNight <= 40 Then
      GILevel = .4125
  ElseIf DayNight <= 60 Then
      GILevel = .325
  ElseIf DayNight <= 80 Then
      GILevel = .2375
  Elseif DayNight <= 100  Then
      GILevel = .15
  End If

  For each xx in GI: xx.IntensityScale = xx.IntensityScale * (GILevel): Next
  For each xx in DTLights: xx.IntensityScale = xx.IntensityScale * (GILevel): Next

End Sub

Sub ALlightsTimer_timer

  if me.uservalue>1 then
      ALlights(me.uservalue-2).state=0
      ALlightsA(me.uservalue-2).state=0
    elseif me.uservalue=1 then
      ALlights(9).state=0
      ALlightsA(9).state=0
    else
      ALlights(8).state=0
      ALlightsA(8).state=0
  end if

  ALlights(me.uservalue).state=1
  ALlightsA(me.uservalue).state=1

  me.uservalue = me.uservalue+1
  if me.uservalue>9 then me.uservalue=0
End sub

Sub FlipperTimer_Timer

'testbox1.text =

' BumperLightA1.state=BumperLight1.state

  LFlip.RotY = LeftFlipper.CurrentAngle
  RFlip.RotY = RightFlipper.CurrentAngle
  LFlip1.RotY = LeftFlipper1.CurrentAngle-90
  RFlip1.RotY = RightFlipper1.CurrentAngle-90
  Pgate.Rotz = Gate.CurrentAngle*0.7

  if sw40.isdropped then
    Lsw40.state=GI3.state
'       l1.state=0
      else
    Lsw40.state=0
'     l1.state=1
    end if

  if sw50.isdropped then
    Lsw50.state=GI3.state
'       l13.state=0
       else
    Lsw50.state=0
'       l13.state=1
    end if

  if sw60.isdropped then
    Lsw60.state=GI3.state
'       l31.state=0
       else
    Lsw60.state=0
'       l31.state=1
    end if

if sw41.isdropped then
    Lsw1.state=GI17.state
'        l6.state=0
      else
    Lsw1.state=0
'   l6.state=1
  end if

  if sw51.isdropped then
    Lsw2.state=GI17.state
'        l7.state=0
    else
    Lsw2.state=0
'   l7.state=1
  end if

  if sw61.isdropped then
    Lsw3.state=GI17.state
'        l8.state=0
    else
    Lsw3.state=0
'   l8.state=1
  end if

    if FlipperShadows=1 then
    FlipperLSh.RotZ = LeftFlipper.currentangle
    FlipperRSh.RotZ = RightFlipper.currentangle
      MiniFlipperLSh.RotZ = LeftFlipper1.currentangle
        MiniFlipperRSh.RotZ = RightFlipper1.currentangle
    end if
End Sub


' Ball locks / kickers

Sub sw46_Hit:
  F13b1.Duration 2, 1000, 0
  F13b2.Duration 2, 1000, 0
  PlaySound "holein"
  PlaySound "Spike bark"
  bslLock.AddBall 0:
End Sub

Sub sw56_Hit:
  PlaySound "holein"
  bsrLock.AddBall 0:
End Sub

Sub SolLLock(enabled)
  If enabled and LeftKickTimer.enabled=0 Then
    bslLock.ExitSol_On
    PkickarmL.RotZ = 15
    LeftKickTimer.uservalue = 0
    LeftKickTimer.Enabled = 1
  End If
End Sub

Sub LeftKickTimer_timer
  select case me.uservalue
    case 5:
    PkickarmL.rotz=0
    me.enabled=0
  end Select
  me.uservalue = me.uservalue+1
End Sub

Sub SolRLock(enabled)
  If enabled and RightKickTimer.enabled=0 Then
   if multiballs<2 then
    multiballs=3

    if l0.state=1 and RightKickTimer.enabled=0 then           'L0 is lit when game enabled
      bMusicOn=false          'turn off autoreplay
      endmusic            'stop regular music
      playmusic "bgout_T&Jmultiballmadness.mp3" 'playmultiball music
      bMusicOn=true         'turn back oon auto play so when multiball music ends music starts again
    end if
   end if

    bsrLock.ExitSol_On
    PkickarmR.RotZ = 15
    RightKickTimer.uservalue = 0
    RightKickTimer.Enabled = 1
  End If
End Sub

Sub RightKickTimer_timer
  select case me.uservalue
    case 5:
    PkickarmR.rotz=0
    me.enabled=0
  end Select
  me.uservalue = me.uservalue+1
End Sub


'******************************************************
'     TROUGH BASED ON NFOZZY'S via bord's pink panther
'******************************************************

Sub sw10_Hit():Controller.Switch(10) = 1:UpdateTrough:End Sub
Sub sw10_UnHit():Controller.Switch(10) = 0:UpdateTrough:End Sub
Sub sw20_Hit():Controller.Switch(20) = 1:UpdateTrough:End Sub
Sub sw20_UnHit():Controller.Switch(20) = 0:UpdateTrough:End Sub
Sub sw76_Hit():Controller.Switch(76) = 1:UpdateTrough:End Sub
Sub sw76_UnHit():Controller.Switch(76) = 0:UpdateTrough:End Sub

Sub UpdateTrough()
  UpdateTroughTimer.Interval = 500
  UpdateTroughTimer.Enabled = 1
End Sub

Sub UpdateTroughTimer_Timer()
  If sw76.BallCntOver = 0 Then sw20.kick 60, 1
  UpdateTroughTimer1.Interval = 500
  UpdateTroughTimer1.Enabled = 1
  Me.Enabled = 0
End Sub

Sub UpdateTroughTimer1_Timer()
  If sw20.BallCntOver = 0 Then sw10.kick 60, 1
  Me.Enabled = 0
End Sub

'******************************************************
'       DRAIN & RELEASE
'******************************************************

Sub sw66_Hit()
  if l0.state=1 then PlaySoundat "drain", sw66
  if multiballs>0 then multiballs=multiballs-1
  if multiballs<2 then              'down to one ball after multiball
    endmusic                  'stop multiball music
    PlayMusic "bgout_T&JbackgroundMusic" & currentsong & ".mp3"                 'play regular musuic
  end if
  startgamesound = 1
    if sw20.BallCntOver > 0 or (sw76.ballcntover>0 and (sw46.ballcntover >0 or sw56.ballcntover>0)) or (sw46.ballcntover>0 and sw56.ballcntover>0) then
    bMusicOn = False  'turn off autoreplay music
    endMusic
  end if
  UpdateTrough
  Controller.Switch(66) = 1
End Sub

Sub sw66_UnHit()
  Controller.Switch(66) = 0
End Sub

Sub solballrelease(enabled)
  If enabled Then
    sw66.kick 60,40
    PlaySoundat SoundFX("fx_Solenoid",DOFContactors), sw66
  End If
End Sub

set Lights(12) = L12

Set LampCallback = GetRef("UpdateMultipleLamps")
Sub UpdateMultipleLamps
    if controller.lamp(12)=true and sw76.BallCntOver > 0 then
    sw76.kick 60, 7

    if bMusicOn = False then
      bMusicOn = TRUE 'TRUE
      playsong        'play random song
    end if

    PlaySoundAt SoundFX("ballrelease",DOFContactors), sw76
    UpdateTrough
    end if

    if controller.lamp(2)=true then dtleft.hit 1:l1.state=0:dtRight.hit 1:l6.state=0
End Sub

 sub startgametimer_timer     '****make random sound for starting game
  Dim v
  if l0.state=1 then
    v = INT(2 * RND(1) )
    Select Case v
      Case 0:PlaySound"introducing"
      Case 1:PlaySound"introducing"
        End Select
  end if
  me.enabled=0
end sub

'**********************************************************************************************************
'Plunger code
'**********************************************************************************************************

Sub TomandJerry_KeyDown(ByVal KeyCode)

    If KeyCode = LeftFlipperKey then FastFlips.FlipL True :  FastFlips.FlipUL True
    If KeyCode = RightFlipperKey then FastFlips.FlipR True :  FastFlips.FlipUR True
  if keycode = addcreditkey then PlaySoundAt "fx_coin", sw66

    If keycode = PlungerKey Then
    Plunger.Pullback
    PlaySoundAt "plungerpull", Plunger
    TimerPlunger.Enabled = True
    TimerPlunger2.Enabled = False
  End If

  If KeyCode = startgamekey and startgametimer.enabled=0 and startgamesound=1 then
    startgamesound=0
    startgametimer.enabled=1
  end if

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

    If KeyDownHandler(keycode) Then Exit Sub

End Sub

Sub TomandJerry_KeyUp(ByVal KeyCode)

    If KeyCode = LeftFlipperKey then FastFlips.FlipL False :  FastFlips.FlipUL False
    If KeyCode = RightFlipperKey then FastFlips.FlipR False :  FastFlips.FlipUR False

    If keycode = PlungerKey Then
    Plunger.Fire
    PlaySoundAt "plunger", Plunger
    TimerPlunger.Enabled = False
    TimerPlunger2.Enabled = True
    PinCab_Shooter.Y = -73.66
  End If


'************************   Start Ball Control 2/3
  if keycode = 203 then bcleft = 0    ' Left Arrow
  if keycode = 200 then bcup = 0      ' Up Arrow
  if keycode = 208 then bcdown = 0    ' Down Arrow
  if keycode = 205 then bcright = 0   ' Right Arrow
'************************   End Ball Control 2/3

    If KeyUpHandler(keycode) Then Exit Sub

End Sub

'************************   Start Ball Control 3/3
Sub StartControl_Hit()
  Set ControlBall = ActiveBall
  contballinplay = true
    End Sub

Sub StopControl_Hit()
  contballinplay = false
End Sub

Dim bcup, bcdown, bcleft, bcright, contball, contballinplay, ControlBall, bcboost
Dim bcvel, bcyveloffset, bcboostmulti

bcboost = 1   'Do Not Change - default setting
bcvel = 4   'Controls the speed of the ball movement
bcyveloffset = -0.01  'Offsets the force of gravity to keep the ball from drifting vertically on the table, should be negative
bcboostmulti = 3  'Boost multiplier to ball veloctiy (toggled with the B key)

Sub BallControl_Timer()
  If Contball and ContBallInPlay then
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



'Drop Targets
 Sub Sw40_Dropped:dtleft.Hit 1 : l1.state=1:StopSound"punch1":PlaySound"punch1": End Sub
 Sub Sw50_Dropped:dtleft.Hit 2 : l13.state=1:StopSound"punch2":PlaySound"punch2": End Sub
 Sub Sw60_Dropped:dtleft.Hit 3 : l31.state=1:StopSound"punch3":PlaySound"punch3": End Sub

 Sub Sw41_Dropped:dtRight.Hit 1: l6.state=0:StopSound"boink1":PlaySound"boink1": End Sub
 Sub Sw51_Dropped:dtRight.Hit 2: l7.state=0:StopSound"boink2":PlaySound"boink2": End Sub
 Sub Sw61_Dropped:dtRight.Hit 3: l8.state=0:StopSound"boink3":PlaySound"boink3": End Sub


 Sub SolRightTargetReset(enabled)
    dim xx
    if enabled then
        dtRight.SolDropUp enabled
    l6.state=2
    l7.state=2
    l8.state=2
    end if
End Sub

Sub SolLeftTargetReset(enabled)
    dim xx
    if enabled then
        dtLeft.SolDropUp enabled
        l1.state=0
    l13.state=0
    l31.state=0
    end if
End Sub

'Bumpers

Sub bumper1_Hit
vpmTimer.PulseSw 30
PlaySoundAt SoundFX("Tomboing",DOFContactors), Bumper1:
DOF 206, DOFPulse
BumperLight1.duration 0, 200, 1
BumperLightA1.duration 0, 200, 1
F13a1.duration 0, 200, 1
F13a2.duration 0, 200, 1
Light2.duration 1, 200, 0
GI15.duration 0, 200, 1
Light4.duration 1, 200, 0
GI21.duration 0, 200, 1
Flasher5a.duration 1, 200, 0
Flasher5a1.duration 1, 200, 0
Flasher5aLIGHTsec.duration 1, 200, 0
Flasher5a1LIGHTsec.duration 1, 200, 0
End Sub

Sub Tomscream_Hit()
if Activeball.vely > 0 then
Playsound "Tomscream1"
end if
end sub

Sub Tomscream2_Hit()
if Activeball.vely > 0 then
Playsound "Tomscream2"
end if
end sub

Sub bumper2_Hit
vpmTimer.PulseSw (60):
PlaySoundAt SoundFX("fx_bumper2",DOFContactors), Bumper1:
DOF 207, DOFPulse
BumperLight2.duration 2, 200, 1
End Sub

Sub bumper3_Hit
vpmTimer.PulseSw (60):
PlaySoundAt SoundFX("fx_bumper3",DOFContactors), Bumper1:
DOF 208, DOFPulse
BumperLight3.duration 2, 200, 1
End Sub

'Wire Triggers

Sub SW38_Hit
al16.Duration 1, 300, 2
End Sub

Sub SW38a_Hit
al16.Duration 0,0,0
end sub

Sub SW39_Hit
al17.Duration 1, 300, 2
End Sub

Sub SW39a_Hit
al17.Duration 0,0,0
end sub

Sub SW42_Hit:Controller.Switch(42)=1:
    if Activeball.vely > 0 then
    Dim w
      w = INT(3 * RND(1) )
      Select Case w
      Case 0:PlaySound"zing3"
      Case 1:PlaySound"Switch2"
      Case 2:PlaySound"Switch2"
            End Select
End if
End Sub
Sub SW42_unHit:Controller.Switch(42)=0:End Sub

Sub SW52_Hit:Controller.Switch(52)=1:
    if Activeball.vely > 0 then
    Dim w
      w = INT(3 * RND(1) )
      Select Case w
      Case 0:PlaySound"zing2"
      Case 1:PlaySound"Switch2"
      Case 2:PlaySound"Switch2"
            End Select
End if
End Sub
Sub SW52_unHit:Controller.Switch(52)=0:End Sub

Sub SW62_Hit:Controller.Switch(62)=1:
    if Activeball.vely > 0 then
    Dim w
      w = INT(3 * RND(1) )
      Select Case w
      Case 0:PlaySound"zing1"
      Case 1:PlaySound"Switch2"
      Case 2:PlaySound"Switch2"
            End Select
End if
End Sub
Sub SW62_unHit:Controller.Switch(62)=0:End Sub

Sub SW73_Hit:Controller.Switch(73)=1:
    PlaySound "Switch"
End Sub
Sub SW73_unHit:Controller.Switch(73)=0:End Sub
Sub SW45_Hit:Controller.Switch(45)=1:
    Playsound "Switch2"
End Sub
Sub SW45_unHit:Controller.Switch(45)=0:End Sub
Sub SW55_Hit:Controller.Switch(55)=1:
  DOF 302, DOFOn
    PlaySound "Switch2"
    al12.Duration 2, 600, 0
    al15.Duration 5, 1000, 0
End Sub
Sub SW55_unHit:Controller.Switch(55)=0:DOF 302, DOFOff:End Sub
Sub SW55a_Hit:Controller.Switch(55)=1:
  DOF 301, DOFOn
    PlaySound "Switch2"
    al13.Duration 2, 600, 0
End Sub
Sub SW55a_unHit:Controller.Switch(55)=0:DOF 301, DOFOff:End Sub
Sub SW65_Hit:Controller.Switch(65)=1:
    PlaySound "Switch2"
    al14.Duration 2, 600, 0
End Sub
Sub SW65_unHit:Controller.Switch(65)=0:End Sub
Sub SW75_Hit:Controller.Switch(75)=1:
    PlaySound "Switch2"
    al11.Duration 2, 600, 0
End Sub
Sub SW75_unHit:Controller.Switch(75)=0:End Sub
Sub SW72_Hit:Controller.Switch(72)=1:End Sub
Sub SW72_unHit:Controller.Switch(72)=0:End Sub
Sub SW70_Hit:Controller.Switch(70)=1:
    PlaySound "Switch2"
End Sub
Sub SW70_unHit:Controller.Switch(70)=0:End Sub
Sub SW74_Hit:Controller.Switch(74)=1:End Sub
Sub SW74_unHit:Controller.Switch(74)=0:End Sub
Sub SW31_Hit:Controller.Switch(31)=1:End Sub
Sub SW31_unHit:Controller.Switch(31)=0:End Sub

Sub SW80_hit
  If Jerry.roty=0 Then
    me.uservalue = 0
    me.timerinterval = 10
    me.timerenabled = 1
  End If
end sub

Sub SW80_timer
  jerry.roty = - me.uservalue * 3
  if jerry.roty < -120 Then
    jerry.roty = -120
    me.timerenabled = 0
  end if
  jerrySHADOW.roty = jerry.roty
  me.uservalue=me.uservalue+1
end sub

Sub SW80c_Hit
  Playsound "Bell"
End Sub

Sub SW81_hit
  If Jerry.roty=-120 Then
    me.uservalue=40
    me.timerinterval = 10
    me.timerenabled= 1
  End If
end sub

Sub SW81_timer
  jerry.roty = - me.uservalue * 3
  if jerry.roty > 0 Then
    jerry.roty = 0
    me.timerenabled = 0
  end if
  jerrySHADOW.roty = jerry.roty
  me.uservalue=me.uservalue - 1
end sub


Sub SW82_hit
  if TomHead.roty=0 and not sw90.timerenabled and not sw93.timerenabled then
    me.uservalue = 0
    me.timerinterval = 10
    me.timerenabled = 1
  End If
end sub

Sub SW82_timer
  TomHead.roty = me.uservalue * 0.75
  if TomHead.roty > 30 Then
    TomHead.roty = 30
    me.timerenabled = 0
  end if
  TomEYES.roty = TomHead.roty
  TomEyeLids.roty = TomHead.roty
  me.uservalue=me.uservalue+1
end sub

Sub SW83_hit
  if TomHead.roty=30 then
    me.uservalue = 40
    me.timerinterval = 10
    me.timerenabled = 1
  End If
end sub

Sub SW83_timer
  TomHead.roty = me.uservalue * 0.75
  if TomHead.roty < 0 Then
    TomHead.roty = 0
    me.timerenabled = 0
  end if
  TomEYES.roty = TomHead.roty
  TomEyeLids.roty = TomHead.roty
  me.uservalue=me.uservalue-1
end sub

Sub SW86_hit
  if TomEyeLids.rotx>6.5 then
    TomEyeLids.rotx = 0
    me.timerinterval = 30
    me.timerenabled = 1
  End If
end sub

Sub SW86_timer
  If TomEyeLids.rotx < 2 Then
    TomEyeLids.rotx = TomEyeLids.rotx + 0.1
  Elseif TomEyeLids.rotx < 6 Then
    TomEyeLids.rotx = TomEyeLids.rotx + 0.2
  Else
    TomEyeLids.rotx = TomEyeLids.rotx + 0.15
    if TomEyeLids.rotx > 7.5 Then
      TomEyeLids.rotx = 7.5
      me.timerenabled = 0
    End If
  End If
  'debug.print TomEyeLids.rotx
end sub

Sub SW88_Hit
  TomHeadLight.Duration 1,1,1
end sub

Sub SW89_Hit
  TomHeadLight.Duration 0,0,0
end sub

Sub SW90_hit
  if TomHead.roty=0  and not sw93.timerenabled and not sw82.timerenabled then
    me.uservalue = 1
    me.timerinterval = 10
    me.timerenabled = 1
  End If
end sub

Sub SW90_timer
  TomHead.roty = TomHead.roty + (me.uservalue * 0.4)
  if TomHead.roty >= 12 Then
    me.uservalue = -1
  end if
  if TomHead.roty < 0 Then
    TomHead.roty = 0
    me.timerenabled = 0
  End If
  TomEyes.roty = TomHead.roty
  TomEyeLids.roty = TomHead.roty
end sub

Sub SW93_hit
  if TomHead.roty=0 and not sw90.timerenabled and not sw82.timerenabled then
    me.uservalue = -1
    me.timerinterval = 10
    me.timerenabled = 1
  End If
end sub

Sub SW93_timer
  TomHead.roty = TomHead.roty + (me.uservalue * 0.6)
  if TomHead.roty <= -18 Then
    me.uservalue = 1
  end if
  if TomHead.roty > 0 Then
    TomHead.roty = 0
    me.timerenabled = 0
  End If
  TomEyes.roty = TomHead.roty
  TomEyeLids.roty = TomHead.roty
end sub

Sub SW95_hit
  Silhouette.transx=0
  me.uservalue = 0
  me.timerinterval = 10
  me.timerenabled = 1
end sub

Sub SW95_timer
  Silhouette.transx = - me.uservalue * 1
  If Silhouette.transx < -170 Then
    me.timerenabled = 0
    Silhouette.transx = 0
  End If
  me.uservalue=me.uservalue+1
end sub

Sub SW95a_Hit
  WindowA.Duration 1,1950,0
  WindowB.Duration 1,1950,0
  WindowC.Duration 1,1950,0
  WindowD.Duration 1,1950,0
end sub

Sub Tomtrough_hit
  if Activeball.vely < 0 then
    me.uservalue=1
    'me.timerenabled= 1
    Dim w
    w = INT(3 * RND(1) )
    Select Case w
      Case 0:PlaySound"Tom Scream Cannon"
      Case 1:PlaySound""
            Case 2:PlaySound""
    End Select
  End if
End Sub

Sub SW96_hit()
  if Activeball.vely < 0 then
    If Trashcan.rotz=0 then
      me.uservalue = 0
      me.timerenabled = 1
      me.timerinterval = 10
      Playsound "garbage can"
    end if
  end if
end sub

Sub SW96_timer
  if me.uservalue > 10 and me.uservalue < 21 then
    Trashcan.rotz = Trashcan.rotz - 1.2
  elseif me.uservalue > 20 and me.uservalue < 31 then
    Trashcan.rotz = Trashcan.rotz + 1.2
  elseif me.uservalue > 30 and me.uservalue < 41 then
    Trashcan.rotz = Trashcan.rotz - 1.2
  elseif me.uservalue > 40 and me.uservalue < 51 then
    Trashcan.rotz = Trashcan.rotz + 1.2
  elseif me.uservalue > 50 and me.uservalue < 61 then
    Trashcan.rotz = Trashcan.rotz - 1.2
  elseif me.uservalue > 60 and me.uservalue < 71 then
    Trashcan.rotz = Trashcan.rotz + 1.2
  elseif me.uservalue > 70 and me.uservalue < 81 then
    Trashcan.rotz = Trashcan.rotz - 1.2
  elseif me.uservalue > 80 and me.uservalue < 91 then
    Trashcan.rotz = Trashcan.rotz + 1.2
  elseif me.uservalue > 90 and me.uservalue < 101 then
    Trashcan.rotz = Trashcan.rotz - 1.2
  elseif me.uservalue > 99 then
    me.timerenabled = 0
    Trashcan.rotz = 0
  End If
  TrashcanSHADOW.rotz = Trashcan.rotz
  TrashcanLid.rotz = Trashcan.rotz
  'debug.print Trashcan.rotz
  me.uservalue=me.uservalue+1
end sub

Sub SW96a_Hit()
    if Activeball.vely < 0 then
        WindowA.Duration 1,3000,0
        WindowB.Duration 1,3000,0
        WindowC.Duration 1,3000,0
        WindowD.Duration 1,3000,0
    end if
end sub

Sub SW96c_hit()
  if Activeball.vely < 0 then
    me.uservalue=1
    'me.timerenabled= 1
    Dim w
    w = INT(3 * RND(1) )
    Select Case w
      Case 0: PlaySound "Thomas get this mouse"
      Case 1: PlaySound "Thomas"
      Case 2: Playsound ""
    End Select
  End if
End Sub


Sub SW98_hit()
  if Activeball.vely > 0 then
    If JerryHAMMER.roty = 0 then
      me.uservalue = 0
      me.timerenabled = 1
      me.timerinterval = 10
      Playsound "Tom inlane left"
    end if
  end if
end sub

Sub SW98_timer
  if me.uservalue < 11 Then
    'do nothing
  elseif me.uservalue < 48 then
    JerryHammer.roty = JerryHammer.roty + 3
  elseif me.uservalue < 68 then
    If me.uservalue = 58 then StarLightsLeft.Duration 1, 1000, 0
  else
    JerryHammer.roty = JerryHammer.roty - 3
    if JerryHammer.roty < 0 Then
      JerryHammer.roty = 0
      me.timerenabled = 0
    end If
  end if
  JerryHAMMERshadow.roty = JerryHammer.roty
  me.uservalue=me.uservalue+1
end sub

Sub SW99_hit()
  if Activeball.vely > 0 then
    If MusclesKnife.roty=0 then
      me.uservalue = 0
      me.timerenabled = 1
      me.timerinterval = 10
    end if
  end if
end sub

Sub SW99_timer
  if me.uservalue < 11 Then
    if me.uservalue = 2 then Playsound "Tom inlane right"
  elseif me.uservalue < 52 then
    MusclesKnife.roty = MusclesKnife.roty - 3.5
  elseif me.uservalue < 72 then
    If me.uservalue = 62 then StarLightsRight.Duration 1, 1000, 0
  else
    MusclesKnife.roty = MusclesKnife.roty + 3.5
    if MusclesKnife.roty > 0 Then
      MusclesKnife.roty = 0
      me.timerenabled = 0
    end If
  end if
  MuscleswithknifeSHADOW.roty = MusclesKnife.roty
  me.uservalue=me.uservalue+1
end sub


Sub SW100_Hit
GI11.Duration 0, 0, 0
end sub

Sub SW101_Hit:Controller.Switch(55)=1:
    PlaySound "ramp"
  if sw46.ballcntover=1 and sw56.ballcntover=1 then
  end if
end sub


Sub SW102_hit
  If Spike.roty=0 and not sw106.timerenabled Then
    me.uservalue=0
    me.timerinterval = 10
    me.timerenabled= 1
  End If
end sub

Sub SW102_timer
  If me.uservalue < 20 Then
    Spike.roty = Spike.roty + 1.5
  elseif me.uservalue < 90 Then
    'do Nothing
  Else
    Spike.roty = Spike.roty - 1.5
    if Spike.roty < 0 Then
      Spike.roty = 0
      me.timerenabled = 0
    end if
  end if
  SpikeSHADOW.roty = Spike.roty
  me.uservalue=me.uservalue + 1
end sub

Sub SW102b_hit
GI17.Duration 2, 1000, 1
GI2.Duration 2, 1000, 1
GI11.Duration 2, 1000, 1
LkickR1.Duration 2, 1000, 1
LkickR.Duration 2, 1000, 1
Lbulb4.Duration 2, 1000, 1
lsw1.Duration 2, 1000, 1
lsw2.Duration 2, 1000, 1
lsw3.Duration 2, 1000, 1
GI35.Duration 2, 1000, 1
GI37.Duration 2, 1000, 1
GI18.Duration 2, 1000, 1
Lbulb11.Duration 2, 1000, 1
Lbulb14.Duration 2, 1000, 1
end sub

Sub SW102c_hit
  if Activeball.vely > 0 then
    me.uservalue=1
    'me.timerenabled= 1
    Dim w
    w = INT(2 * RND(1) )
    Select Case w
      Case 0:PlaySound"That's my boy"
      Case 1:PlaySound"Listen pussycat"
    End Select
  End if
End Sub

Sub SW103_Hit
  GI9.Duration 2, 900, 1
    PlaySound "drain1"
end sub

Sub leftoutlane_Hit()
  if Activeball.vely > 0 then
    me.uservalue=1
    me.timerenabled= 1
  end if
end sub

Sub leftoutlane_timer
select Case me.uservalue
Case 1: Playsound "Bomb fuse"
Case 2: Playsound ""
Case 3: Playsound ""
Case 4: Playsound ""
Case 5: Playsound ""
Case 6: Playsound ""
Case 7: Playsound ""
Case 8: Playsound ""
Case 9: Playsound ""
Case 10: Stopsound "Bomb fuse"
Case 11: Playsound "Bomb left"
Case 12: GI32.Duration 2, 1050, 0: me.timerenabled=0
end select
me.uservalue=me.uservalue+1
end sub

Sub rightoutlane_Hit()
  if Activeball.vely > 0 then
    me.uservalue=1
    me.timerenabled= 1
  end if
end sub

Sub rightoutlane_timer
select Case me.uservalue
Case 1: Playsound "Bomb fuse"
Case 2: Playsound ""
Case 3: Playsound ""
Case 4: Playsound ""
Case 5: Playsound ""
Case 6: Playsound ""
Case 7: Playsound ""
Case 8: Playsound ""
Case 9: Playsound ""
Case 10: Stopsound "Bomb fuse"
Case 11: Playsound "Bomb right"
Case 12: GI32.Duration 2, 1000, 0: me.timerenabled=0
end select
me.uservalue=me.uservalue+1
end sub

Sub SW104_Hit
Llsling2.Duration 2, 500, 1
end sub


Sub SW105_hit
  If Tuffy.transx = 0 then
    F13a.Duration 0, 2900, 0
    al6a.Duration 0, 2900, 0
    F13.Duration 0, 2900, 0
    me.uservalue=1
    me.timerenabled= 1
    me.timerinterval = 10
  End If
end sub

Sub SW105_timer
  if me.uservalue < 10 then
    Tuffy.transx = Tuffy.transx - 1
  elseif me.uservalue < 40 then
    Tuffy.transx = Tuffy.transx - 1.5
  elseif me.uservalue < 50 then
    Tuffy.transx = Tuffy.transx - 1.3
  elseif me.uservalue < 420 then
    if me.uservalue = 60 then PlaySound "Touche Pussycat"
    if me.uservalue = 70 then TuffyLight.Duration 1, 3300, 0
    if me.uservalue = 80 then TomHeadLight.Duration 1, 2000, 0
  elseif me.uservalue < 430 then
    Tuffy.transx = Tuffy.transx + 1.3
  elseif me.uservalue < 460 then
    Tuffy.transx = Tuffy.transx + 1.5
  else
    Tuffy.transx = Tuffy.transx + 1
    if Tuffy.transx > 0 Then
      Tuffy.transx = 0
      me.timerenabled = 0
    End If
  End If
  TuffyShadow.transx = Tuffy.transx
  me.uservalue=me.uservalue+1
end sub

Sub SW106_hit
  If Spike.roty=0 and not sw102.timerenabled Then
    me.uservalue=0
    me.timerinterval = 10
    me.timerenabled= 1
  End If
end sub

Sub SW106_timer
  If me.uservalue < 20 Then
    if me.uservalue = 0 then GI37.Duration 2, 1500, 0
    if me.uservalue = 10 then Lbulb11.Duration 2, 1500, 0
  elseif me.uservalue < 60 Then
    Spike.roty = Spike.roty + 2
  elseif me.uservalue < 140 Then
    'DO NOTHING
  Else
    Spike.roty = Spike.roty - 2
    if Spike.roty < 0 Then
      Spike.roty = 0
      me.timerenabled = 0
    end if
  end if
  SpikeSHADOW.roty = Spike.roty
  me.uservalue=me.uservalue + 1
end sub


Sub SW107_hit()
  if Activeball.vely > 0 then
    if Tyke.rotz=0 then
      me.uservalue = -1
      me.timerinterval = 10
      me.timerenabled = 1
    end if
  end if
end sub

Sub SW107_timer
  Tyke.rotz = Tyke.rotz + (me.uservalue * 0.9)
  if Tyke.rotz < - 26.5 Then
    me.uservalue = 1
  end If
  if tyke.rotz > 0 Then
    tyke.rotz = 0
    me.timerenabled = 0
  end if
  TykeSHADOW.rotz = Tyke.rotz
end sub

Sub SWTomsplat_hit
if Activeball.vely > 0 then
me.uservalue=1
'me.timerenabled= 1
Dim w
      w = INT(3 * RND(1) )
      Select Case w
      Case 0:PlaySound"Tom splat"
      Case 1:PlaySound""
      Case 2:PlaySound""
End Select
End if
End Sub

Sub SWTykebark_hit
if Activeball.vely > 0 then
me.uservalue=1
'me.timerenabled= 1
Dim w
      w = INT(3 * RND(1) )
      Select Case w
      Case 0:PlaySound"Tyke bark"
      Case 1:PlaySound""
      Case 2:PlaySound""
End Select
End if
End Sub

Sub SW108_Hit
    if Activeball.vely < 0 then
        Flasher5a.Duration 1, 700, 0
        Flasher5aLIGHTsec.duration 1, 700, 0
        Playsound "Ramp"
        me.uservalue=1
    me.timerenabled= 1
  end if
end sub

Sub SW109_Hit
Flasher5a1.Duration 1, 700, 0
Flasher5a1LIGHTsec.Duration 1, 700, 0
end sub

Sub JerryFuseOn_Hit
  if Activeball.vely > 0 then
    Playsound "Cannon fuse"
        Fuse.Duration 1, 1, 2
        FuseReflect1.Duration 1, 1, 2
        FuseReflect2.Duration 1, 1, 2
    me.uservalue=1
    me.timerenabled= 1
  end if
end sub

Sub JerryFuseOff_Hit
  if Activeball.vely < 0 then
    Stopsound "Cannon fuse"
        Playsound "Cannon fire"
        Fuse.Duration 0, 0, 0
        FuseReflect1.Duration 0, 0, 0
        FuseReflect2.Duration 0, 0, 0
    me.uservalue=1
    me.timerenabled= 1
  end if
end sub

Sub Cannon_Hit()
  if Activeball.vely < 0 then
    me.uservalue=1
    me.timerenabled= 1
  end if
end sub

Sub Cannon_timer
select Case me.uservalue
Case 1: CannonFire.Duration 2, 500, 0
Case 2: CannonFireReflect.Duration 2, 500, 0
Case 3: TomTroughLight.Duration 1, 750, 0: me.timerenabled=0
end select
me.uservalue=me.uservalue+1
end sub

'Targets
Sub sw43_Hit:vpmTimer.PulseSw (43):
    PlaySound "target1"
End Sub
Sub sw53_Hit:vpmTimer.PulseSw (53):
    PlaySound "target2"
End Sub
Sub sw63_Hit:vpmTimer.PulseSw (63):
    PlaySound "target3"
End Sub
Sub sw44_Hit:vpmTimer.PulseSw (44):
    PlaySound "target1"
End Sub
Sub sw54_Hit:vpmTimer.PulseSw (54):
    PlaySound "target2"
End Sub
Sub sw64_Hit:vpmTimer.PulseSw (64):
    PlaySound "Jerry gulp"
End Sub
Sub sw73a_Hit:vpmTimer.PulseSw (73):
    PlaySound "Jerry gulp":
  DOF 301, DOFPulse
End Sub



Sub SolKnocker(Enabled)
    If Enabled Then PlaySound SoundFX("Knocker",DOFKnocker)
End Sub


'*****************************************
' Ramp and Drop Sounds
'*****************************************

Sub RedRampStart_Hit()
  if Activeball.vely>0 then
    PlaySoundAtBall "plasticroll"
       else
    PlaySoundAtBall "plasticroll"
        PlaySoundAtBall ""
      end if
End Sub

Sub RedRampStop_Hit()
  StopSound "plasticroll"
  StopSound "plasticroll"
    StopSound "ramp"
End Sub

Sub RedRampStop1_Hit()
  StopSound "plasticroll"
  StopSound "plasticroll"
  vpmTimer.AddTimer 200, "BallDropSound(RedRampStop1)'"
End Sub


Sub BallDropSound(loc)
  PlaySoundAt "BallDrop", loc
End Sub


Sub HotShotStart_Hit()
  If Activeball.vely < 0 Then
    PlaySoundAtBall "plasticroll"
  End If
End Sub

Sub HotShotStart_Unhit()
  If Activeball.vely > 0 Then
    StopSound "plasticroll"
  End If
End Sub

Sub HotShotStop_UnHit()
  If Activeball.vely < 0 Then
    GI1.Duration 2, 500, 1
        StopSound "plasticroll"
    vpmTimer.AddTimer 200, "BallDropSound(HotShotStop)'"
  End If
End Sub

'**********Rubber Animations

sub RslingA_hit
  SlingA.visible=0
  SlingA1.visible=1
  me.uservalue=1
  Me.timerenabled=1
end sub

sub RslingA_timer                 'default 50 timer
  select case me.uservalue
    Case 1: SlingA1.visible=0: SlingA.visible=1
    case 2: SlingA.visible=0: SlingA2.visible=1
    Case 3: SlingA2.visible=0: SlingA.visible=1: Me.timerenabled=0
  end Select
  me.uservalue=me.uservalue+1
end sub


'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep, Tstep, URstep


Sub RightSlingShot_Slingshot
    PlaySoundAt SoundFX("fx_slingshot",DOFContactors), slingR
    DOF 202, DOFPulse
    vpmtimer.PulseSw(32)
    RSling.Visible = 0
    RSling1.Visible = 1
    slingR.objroty = -15
    RStep = 0
    RightSlingShot.TimerEnabled = 1
    GI6.Duration 2, 300, 1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:slingR.objroty = -7
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:slingR.objroty = 0:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
    PlaySoundAt SoundFX("fx_slingshot",DOFContactors), slingL
    DOF 201, DOFPulse
    vpmtimer.pulsesw(32)
    LSling.Visible = 0
    LSling1.Visible = 1
    slingL.objroty = 15
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
    GI7.Duration 2, 300, 1
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:slingL.objroty = 7
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:slingL.objroty=0:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

Sub TopSlingShot_Slingshot
    PlaySoundAt SoundFX("fx_slingshot",DOFContactors), slingT
    DOF 203, DOFPulse
    vpmtimer.pulsesw(32)
    Tsling.Visible = 0
    Tsling1.Visible = 1
    slingT.objroty = 15
    TStep = 0
    TopSlingShot.TimerEnabled = 1
End Sub

Sub TopSlingShot_Timer
    Select Case TStep
        Case 3:Tsling1.Visible = 0:Tsling2.Visible = 1:slingT.objroty = 7
        Case 4:Tsling2.Visible = 0:Tsling.Visible = 1:slingT.objroty=0:TopSlingShot.TimerEnabled = 0
    End Select
    TStep = TStep + 1
End Sub

Sub URightSlingShot_Slingshot
    PlaySoundAt SoundFX("fx_slingshot",DOFContactors), slingUR
    DOF 204, DOFPulse
    vpmtimer.PulseSw(32)
    URSling.Visible = 0
    URSling1.Visible = 1
    slingUR.objroty = -15
    URStep = 0
    URightSlingShot.TimerEnabled = 1
    Playsound "Wham"
    GI5.Duration 2, 300, 1
End Sub

Sub URightSlingShot_Timer
    Select Case URStep
        Case 3:URSLing1.Visible = 0:URSLing2.Visible = 1:slingUR.objroty = -7
        Case 4:URSLing2.Visible = 0:URSLing.Visible = 1:slingUR.objroty = 0:URightSlingShot.TimerEnabled = 0
    End Select
    URStep = URStep + 1
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


'*********************************************************************
'                     Supporting Ball & Sound Functions
'*********************************************************************

Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "table1" is the name of the table
  Dim tmp
    tmp = tableobj.y * 2 / TomandJerry.height-1
    If tmp > 0 Then
    AudioFade = Csng(tmp ^10)
    Else
        AudioFade = Csng(-((- tmp) ^10) )
    End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "table1" is the name of the table
    Dim tmp
    tmp = tableobj.x * 2 / TomandJerry.width-1
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
'      Ball Rolling Sounds by JP
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
        If BallVel(BOT(b) ) > 1 AND BOT(b).z < 30 Then
            rolling(b) = True
            PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b)), AudioPan(BOT(b)), 0, Pitch(BOT(b)), 1, 0, AudioFade(BOT(b))
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

'**********************
' Object Hit Sounds
'**********************
Sub a_Woods_Hit(idx)
  PlaySound "fx_Woodhit", 0, Vol(ActiveBall), Audiopan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub a_Triggers_Hit(idx)
  PlaySound "fx_switch", 0, Vol(ActiveBall), Audiopan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub a_Pins_Hit (idx)
  PlaySound "pinhit_low", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub a_Targets_Hit (idx)
  PlaySound SoundFX("fx_target",DOFTargets), 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub a_DropTargets_Hit (idx)
  PlaySound "fx_drop1", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub a_Metals_Thin_Hit (idx)
  PlaySound "metalhit_thin", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub a_Metals_Medium_Hit (idx)
  PlaySound "metalhit_medium", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub a_Metals2_Hit (idx)
  PlaySound "metalhit2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub a_Gates_Hit (idx)
  PlaySound "gate4", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Spinner1_Spin
PlaySound "fx_spinner", 0, .25, AudioPan(Spinner1), 0.25, 0, 0, 1, AudioFade(Spinner1):DOF 303, DOFPulse
End Sub

Sub a_Rubbers_Hit(idx)
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 20 then
    PlaySound "fx_rubber2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End if
  If finalspeed >= 6 AND finalspeed <= 20 then
    RandomSoundRubber()
  End If
End Sub

Sub a_Posts_Hit(idx)
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


'*********************************
'GI uselamps workaround by nfozzy
'*********************************

dim GIlamps : set GIlamps = New GIcatcherobject
Class GIcatcherObject   'object that disguises itself as a light. (UseLamps workaround for System80 GI circuit)
    Public Property Let State(input)
        dim x
        if input = 1 then 'If GI switch is engaged, turn off GI.
            for each x in gi : x.state = 0 : next
        elseif input = 0 then
            for each x in gi : x.state = 1 : next
        end if
        'tb.text = "gitcatcher.state = " & input    'debug
    End Property
End Class


set Lights(1) = GIlamps 'GI circuit


'*****************************************
'           BALL SHADOW by ninnuzu
'*****************************************
Dim BallShadow
BallShadow = Array (BallShadow1,BallShadow2,BallShadow3,BallShadow4,BallShadow5)

Sub BallShadowUpdate_timer()
    Dim BOT, b
  Dim maxXoffset
  maxXoffset=6
    BOT = GetBalls

  ' render the shadow for each ball
    For b = 0 to UBound(BOT)
    BallShadow(b).X = BOT(b).X-maxXoffset*(1-(Bot(1).X)/(TomandJerry.Width/2))
    BallShadow(b).Y = BOT(b).Y + 12
    If BOT(b).Z > 0 and BOT(b).Z < 30 Then
      BallShadow(b).visible = 1
    Else
      BallShadow(b).visible = 0
    End If
  Next
End Sub

'Gottlieb Hollywood Heat
'added by Inkochnito
Sub editDips
  Dim vpmDips : Set vpmDips = New cvpmDips
  With vpmDips

    .AddForm 700,400,"Tom & Jerry - DIP Switches"
    .AddFrame 2,4,190,"Maximum Credits",49152,Array("8 Credits",0,"10 Credits",32768,"15 Credits",&H00004000,"20 Credits",49152)'dip 15&16
    .AddFrame 2,80,190,"Coin Chute 1 and 2 Control",&H00002000,Array("Separate",0,"Same",&H00002000)'dip 14
    .AddFrame 2,126,190,"Playfield Special",&H00200000,Array("Replay",0,"Extra Ball",&H00200000)'dip 22
    .AddFrame 2,172,190,"High Games to Date Control",&H00000020,Array("No Effect",0,"Reset High Games 2-5 on Power Off",&H00000020)'dip6
    .AddFrame 2,218,190,"Added a Letter to CARTOON when",&H40000000,Array("Top Rollovers are Completed",0,"Top Rollovers are Lit",&H40000000)'dip 31
    .AddFrame 2,264,190,"Drop Target Lamps",&H80000000,Array("Reset Ball to Ball",0,"Memorize Ball to Ball",&H80000000)'dip 32
    .AddFrame 205,4,190,"High Score to Date Awards",&H00C00000,Array("Not Displayed and No Award",0,"Displayed and No Award",&H00800000,"Displayed and 2 Replays",&H00400000,"Displayed and 3 Replays",&H00C00000)'dip 23&24
    '.AddFrame 205,80,190,"Balls per game",&H01000000,Array("5 balls",0,"3 balls",&H01000000)'dip 25
    .AddFrame 205,126,190,"Replay Limit",&H04000000,Array("No Limit",0,"One per Ball",&H04000000)'dip 27
    .AddFrame 205,172,190,"Novelty",&H08000000,Array("Normal",0,"Extra Ball and Replay Score 500000",&H08000000)'dip 28
    .AddFrame 205,218,190,"Game Mode",&H10000000,Array("Replay",0,"Extra Ball",&H10000000)'dip 29
    .AddFrame 205,264,190,"3rd Coin Chute Credits Control",&H20000000,Array("No Effect",0,"Add 9",&H20000000)'dip 30
    .AddChk 205,316,180,Array("Match Feature",&H02000000)'dip 26
    '.AddChk 2,316,190,Array("Attract Sound",&H00000040)'dip 7
    .AddLabel 50,340,300,20,"After Hitting OK, Press F3 to Reset Game with New Settings."
    .ViewDips
  End With

End Sub
Set vpmShowDips = GetRef("editDips")

'Finding an individual dip state based on scapino's Strikes and spares dip code - from unclewillys pinball pool


 Sub DipsTimer_Timer()
  Dim TheDips(32)
  Dim BPG
    Dim DipsNumber

  DipsNumber = Controller.Dip(3)

  TheDips(32) = Int(DipsNumber/128)
  If TheDips(32) = 1 then DipsNumber = DipsNumber - 128 end if
  TheDips(31) = Int(DipsNumber/64)
  If TheDips(31) = 1 then DipsNumber = DipsNumber - 64 end if
  TheDips(30) = Int(DipsNumber/32)
  If TheDips(30) = 1 then DipsNumber = DipsNumber - 32 end if
  TheDips(29) = Int(DipsNumber/16)
  If TheDips(29) = 1 then DipsNumber = DipsNumber - 16 end if
  TheDips(28) = Int(DipsNumber/8)
  If TheDips(28) = 1 then DipsNumber = DipsNumber - 8 end if
  TheDips(27) = Int(DipsNumber/4)
  If TheDips(27) = 1 then DipsNumber = DipsNumber - 4 end if
  TheDips(26) = Int(DipsNumber/2)
  If TheDips(26) = 1 then DipsNumber = DipsNumber - 2 end if
  TheDips(25) = Int(DipsNumber)

  BPG = TheDips(25)
  If BPG = 1 then
'   instcard.image="InstCard3Balls"
    Else
'   instcard.image="InstCard5Balls"
  End if
  DipsTimer.enabled=0
 End Sub

Sub TomandJerry_Paused:Controller.Pause = 1:End Sub
Sub TomandJerry_unPaused:Controller.Pause = 0:End Sub

Sub TomandJerry_Exit
  Controller.Games(cGameName).Settings.Value("sound")=1
  Controller.Stop
End Sub


'cFastFlips by nFozzy
'Bypasses pinmame callback for faster and more responsive flippers
'Version 1.1 beta2 (More proper behaviour, extra safety against script errors)
'*************************************************
Function NullFunction(aEnabled):End Function    '1 argument null function placeholder
Class cFastFlips
    Public TiltObjects, DebugOn, hi
    Private SubL, SubUL, SubR, SubUR, FlippersEnabled, Delay, LagCompensation, Name, FlipState(3)

    Private Sub Class_Initialize()
        Delay = 0 : FlippersEnabled = False : DebugOn = False : LagCompensation = False
        Set SubL = GetRef("NullFunction"): Set SubR = GetRef("NullFunction") : Set SubUL = GetRef("NullFunction"): Set SubUR = GetRef("NullFunction")
    End Sub

    'set callbacks
    Public Property Let CallBackL(aInput)  : Set SubL  = GetRef(aInput) : Decouple sLLFlipper, aInput: End Property
    Public Property Let CallBackUL(aInput) : Set SubUL = GetRef(aInput) : End Property
    Public Property Let CallBackR(aInput)  : Set SubR  = GetRef(aInput) : Decouple sLRFlipper, aInput:  End Property
    Public Property Let CallBackUR(aInput) : Set SubUR = GetRef(aInput) : End Property
    Public Sub InitDelay(aName, aDelay) : Name = aName : delay = aDelay : End Sub   'Create Delay
    'Automatically decouple flipper solcallback script lines (only if both are pointing to the same sub) thanks gtxjoe
    Private Sub Decouple(aSolType, aInput)  : If StrComp(SolCallback(aSolType),aInput,1) = 0 then SolCallback(aSolType) = Empty End If : End Sub

    'call callbacks
    Public Sub FlipL(aEnabled)
        FlipState(0) = aEnabled 'track flipper button states: the game-on sol flips immediately if the button is held down (1.1)
        If not FlippersEnabled and not DebugOn then Exit Sub
        subL aEnabled
    End Sub

    Public Sub FlipR(aEnabled)
        FlipState(1) = aEnabled
        If not FlippersEnabled and not DebugOn then Exit Sub
        subR aEnabled
    End Sub

    Public Sub FlipUL(aEnabled)
        FlipState(2) = aEnabled
        If not FlippersEnabled and not DebugOn then Exit Sub
        subUL aEnabled
    End Sub

    Public Sub FlipUR(aEnabled)
        FlipState(3) = aEnabled
        If not FlippersEnabled and not DebugOn then Exit Sub
        subUR aEnabled
    End Sub

    Public Sub TiltSol(aEnabled)    'Handle solenoid / Delay (if delayinit)
        If delay > 0 and not aEnabled then  'handle delay
            vpmtimer.addtimer Delay, Name & ".FireDelay" & "'"
            LagCompensation = True
        else
            If Delay > 0 then LagCompensation = False
            EnableFlippers(aEnabled)
        end If
    End Sub

    Sub FireDelay() : If LagCompensation then EnableFlippers False End If : End Sub

    Private Sub EnableFlippers(aEnabled)
        If aEnabled then SubL FlipState(0) : SubR FlipState(1) : subUL FlipState(2) : subUR FlipState(3)
        FlippersEnabled = aEnabled
        If TiltObjects then vpmnudge.solgameon aEnabled
        If Not aEnabled then
            subL False
            subR False
            If not IsEmpty(subUL) then subUL False
            If not IsEmpty(subUR) then subUR False
        End If
    End Sub


End Class

' VR PLUNGER ANIMATION
'
' Code needed to animate the plunger. If you pull the plunger it will move in VR.
' IMPORTANT: there are two numeric values in the code that define the postion of the plunger and the
' range in which it can move. The first numeric value is the actual y position of the plunger primitive
' and the second is the actual y position + 100 to determine the range in which it can move.
'
' You need to to select the VR_Primary_plunger primitive you copied from the
' template and copy the value of the Y position
' (e.g. 2130) into the code. The value that determines the range of the plunger is always the y
' position + 100 (e.g. 2230).
'

Sub TimerPlunger_Timer

  If PinCab_Shooter.Y < 26.34 then
    PinCab_Shooter.Y = PinCab_Shooter.Y + 5
  End If

End Sub

Sub TimerPlunger2_Timer
  PinCab_Shooter.Y = -73.66 + (5* Plunger.Position) -20
End Sub

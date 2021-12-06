Option Explicit
Randomize

Const cGameName = "arena"

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "Can't open controller.vbs"
On Error Goto 0

LoadVPM "01210000", "sys80.VBS", 3.1

' Thalamus 2020 January : Improved directional sounds
' !! NOTE : Table not verified yet !!

' Options
' Volume devided by - lower gets higher sound

Const VolDiv = 2000    ' Lower number, louder ballrolling/collition sound
Const VolCol = 10      ' Ball collition divider ( voldiv/volcol )

'**********************************************************
'********       OPTIONS     *******************************
'**********************************************************

Dim BallShadows: Ballshadows=0          '******************set to 1 to turn on Ball shadows
Dim FlipperShadows: FlipperShadows=1  '***********set to 1 to turn on Flipper shadows
Dim ROMSounds: ROMSounds=1        '**********set to 0 for no rom sounds, 1 to play rom sounds.. mostly used for testing


'************************************************
'************************************************
'************************************************
'************************************************
'************************************************
Const UseSolenoids = 2
Const UseLamps = True
Const UseSync = False
Const UseGI = False

' Standard Sounds
Const SSolenoidOn = "fx_solenoid"
Const SSolenoidOff = "fx_solenoidoff"
Const SCoin = "fx_coin"

Dim bsTrough, bslLock, bsRLock, FastFlips, objekt, xx
Dim showdmd


Sub Arena_Init
  if B2SOn then
    showdmd=0
    else
    showdmd=1
  end if
     With Controller
         .GameName = cGameName
         If Err Then MsgBox "Can't start Game " & cGameName & vbNewLine & Err.Description:Exit Sub
         .SplashInfoLine = "Alien Star (Gottlieb 1984)"&chr(13)&"1.0"
         .HandleKeyboard = 0
         .ShowTitle = 0
         .ShowDMDOnly = showdmd
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
        .InitSaucer sw21,21, 120,8.5
        .InitExitSnd Soundfx("holekick",DOFContactors), Soundfx("HoleKick",DOFContactors)
    end with

    Set bsRLock=New cvpmBallStack
    with bsRLock
        .InitSaucer sw31,31,120,8.5
        .InitExitSnd Soundfx("holekick",DOFContactors), Soundfx("HoleKick",DOFContactors)
    end with

' Nudging
  vpmNudge.TiltSwitch = 57
    vpmNudge.Sensitivity = 1
    vpmNudge.TiltObj = Array(leftslingshot, rightslingshot)


  if ballshadows=1 then
        BallShadowUpdate.enabled=1
      else
        BallShadowUpdate.enabled=0
    end if

    if flippershadows=1 then
        FlipperLSh.visible=1
        FlipperRSh.visible=1
       else
        FlipperLSh.visible=0
        FlipperRSh.visible=0

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

End Sub



'************************************************
' Solenoids
'************************************************
SolCallback(2) =    "SolBankReset"
SolCallback(4) =    "SolFlasherR"
SolCallback(5) =    "SolRLock"
SolCallback(6) =    "SolLLock"
SolCallback(7) =    "SolFlasherL"
SolCallback(8) =    "solknocker"
SolCallback(9) =    "solballrelease"
SolCallback(10) = "FastFlips.TiltSol"

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
     If Enabled Then
        PlaySoundAtVol SoundFX("fx_Flipperup",DOFFlippers), LeftFlipper, 1:LeftFlipper.RotateToEnd
'   controller.Switch(48)=1
     Else
        PlaySoundAtVol SoundFX("fx_Flipperdown",DOFFlippers), LeftFlipper, 1:LeftFlipper.RotateToStart
'   controller.Switch(48)=0
     End If
  End Sub


Sub SolRFlipper(Enabled)
     If Enabled Then
         PlaySoundAtVol SoundFX("fx_Flipperup",DOFFlippers), RightFlipper, 1:RightFlipper.RotateToEnd
    controller.Switch(45)=1
     Else
         PlaySoundAtVol SoundFX("fx_Flipperdown",DOFFlippers), RightFlipper, 1:RightFlipper.RotateToStart
    controller.Switch(45)=0
     End If
End Sub

Sub SolFlasherR(enabled)
  if enabled Then
    F4.state=1
    F4a.state=1
    else
    F4.state=0
    F4a.state=0
  end if
end sub

Sub SolFlasherL(enabled)
  if enabled Then
    F7.state=1
    F7a.state=1
    else
    F7.state=0
    F7a.state=0
  end if
end sub

Sub SolBankReset(Enabled)
  If Enabled Then
    PlaySoundAt "DTreset", sw40
    for each objekt in DropTargets: objekt.isdropped=0: Next
  end if
end sub

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

End Sub

Sub DSLightsTimer_timer

  DSLights(DSLightsTimer.uservalue).duration 1, 40, 0
  DSLightsTimer.uservalue = DSLightsTimer.uservalue + 1
  if DSLightsTimer.uservalue>9 then DSLightsTimer.uservalue=0

end sub

Sub FlipperTimer_Timer
  Dim PI: PI=3.1415926

'testbox.text = "LDS1.timer "&DSLightsTimer.enabled
'testbox1.text = "LDS1.uv "&DSLightsTimer.uservalue
'testbox2.text = ((DSLightsTimer.uservalue+3) MOD 9)

  LFlip.RotZ = LeftFlipper.CurrentAngle
  RFlip.RotZ = RightFlipper.CurrentAngle
  Lflip1.RotZ = LeftFlipper.CurrentAngle
  RFlip1.RotZ = RightFlipper.CurrentAngle
' Pgate.rotx = Gate.currentangle*0.6
' Pgate1.rotx = Gate1.currentangle*0.6

  Dim SpinnerRadius: SpinnerRadius=7

  SpinnerRod.TransZ = (cos((sw24.CurrentAngle + 180) * (PI/180))+1) * SpinnerRadius
  SpinnerRod.TransY = sin((sw24.CurrentAngle) * (PI/180)) * -SpinnerRadius


  if FlipperShadows=1 then
    FlipperLSh.RotZ = LeftFlipper.currentangle
    FlipperRSh.RotZ = RightFlipper.currentangle
  end if
End Sub


' Ball locks / kickers

Sub sw21_Hit:PlaySoundAtVol "holein", ActiveBall, 1:bslLock.AddBall 0:End Sub
Sub sw31_Hit:PlaySoundAtVol "holein", ActiveBall, 1:bsRLock.AddBall 0:End Sub


Sub SolLLock(enabled)
  If enabled Then
    bslLock.ExitSol_On
    DOF 106, DOFpulse
    LeftKickTimer.uservalue = 0
    PkickarmL.RotZ = 15
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
  If enabled Then
    bsRLock.ExitSol_On
    DOF 106, DOFpulse
    RightKickTimer.uservalue = 0
    PkickarmR.RotZ = 15
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
  If sw76.BallCntOver = 0 Then sw20.kick 54, 1
  UpdateTroughTimer1.Interval = 500
  UpdateTroughTimer1.Enabled = 1
  Me.Enabled = 0
End Sub

Sub UpdateTroughTimer1_Timer()
  If sw20.BallCntOver = 0 Then sw10.kick 54, 1
  Me.Enabled = 0
End Sub

'******************************************************
'       DRAIN & RELEASE
'******************************************************

Sub sw66_Hit()
  PlaySoundat "drain", sw66
  UpdateTrough
  Controller.Switch(66) = 1
End Sub

Sub sw66_UnHit()
  Controller.Switch(66) = 0
End Sub

Sub solballrelease(enabled)
  If enabled Then
    sw66.kick 54,40
    PlaySoundat SoundFX("fx_Solenoid",DOFContactors), sw66
  End If
End Sub

set Lights(12) = L12

Set LampCallback = GetRef("UpdateMultipleLamps")
Sub UpdateMultipleLamps
    if controller.lamp(12)=true and sw76.BallCntOver > 0 then
    sw76.kick 60, 7
    PlaySoundAt SoundFX("ballrelease",DOFContactors), sw76
    UpdateTrough
    end if

  if controller.lamp(13)=true then    'Inside Gate Relay
    InsideGate.rotatetoend
    controller.Switch(35)=1
    else
    InsideGate.rotatetostart
    Controller.Switch(35)=0
  end if

  if controller.lamp(14)=true then    'Outside Gate Relay
    OutsideGate.rotatetoend
    else
    OutsideGate.rotatetostart
  end if

  if controller.lamp(17)=true then    'ramp DS lights
    DSLightsTimer.enabled=1
    else
    DSLightsTimer.enabled=0
    for each objekt in DSLights: objekt.state=0: next
  end if

  if Controller.Lamp(24) then
    F24.visible = True
    else
    F24.visible = False
  end if

  if Controller.Lamp(25) then
    F25.visible = True
    else
    F25.visible = False
  end if

  if Controller.Lamp(26) then
    F26.visible = True
    F26A.visible = True
    else
    F26.visible = False
    F26A.visible = False
  end if

End Sub



'**********************************************************************************************************
'Plunger code
'**********************************************************************************************************

Sub Arena_KeyDown(ByVal KeyCode)
    If KeyCode = LeftFlipperKey then FastFlips.FlipL True :  FastFlips.FlipUL True
    If KeyCode = RightFlipperKey then FastFlips.FlipR True :  FastFlips.FlipUR True

    If keycode = PlungerKey Then Plunger.Pullback:PlaySoundAt "plungerpull", Plunger

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

Sub Arena_KeyUp(ByVal KeyCode)
    If KeyCode = LeftFlipperKey then FastFlips.FlipL False :  FastFlips.FlipUL False
    If KeyCode = RightFlipperKey then FastFlips.FlipR False :  FastFlips.FlipUR False

    If keycode = PlungerKey Then
    Plunger.Fire
    if startcontrol.BallCntOver>0 then
      PlaySoundAt "plungerreleaseball", Plunger
      else
      PlaySoundAt "plungerreleasefree", Plunger
    end if
  end if

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



'Wire Triggers


Sub SW22_Hit:Controller.Switch(22)=1:End Sub
Sub SW22_unHit:Controller.Switch(22)=0:End Sub
Sub SW30_Hit:Controller.Switch(30)=1:End Sub
Sub SW30_unHit:Controller.Switch(30)=0:End Sub
Sub SW32_Hit:Controller.Switch(32)=1:End Sub
Sub SW32_unHit:Controller.Switch(32)=0:End Sub
Sub SW42_Hit:Controller.Switch(42)=1:End Sub
Sub SW42_unHit:Controller.Switch(42)=0:End Sub

Sub SW44_Hit:Controller.Switch(44)=1:End Sub
Sub SW44_unHit:Controller.Switch(44)=0:End Sub

Sub SW54_Hit:Controller.Switch(54)=1:End Sub
Sub SW54_unHit:Controller.Switch(54)=0:End Sub
Sub SW55_Hit:Controller.Switch(55)=1:End Sub
Sub SW55_unHit:Controller.Switch(55)=0:End Sub

Sub SW41_Hit:Controller.Switch(41)=1:End Sub
Sub SW41_unHit:Controller.Switch(41)=0:End Sub
Sub SW51_Hit:Controller.Switch(51)=1:End Sub
Sub SW51_unHit:Controller.Switch(51)=0:End Sub
Sub SW61_Hit:Controller.Switch(61)=1:End Sub
Sub SW61_unHit:Controller.Switch(61)=0:End Sub
Sub SW71_Hit:Controller.Switch(71)=1:End Sub
Sub SW71_unHit:Controller.Switch(71)=0:End Sub

Sub SW64_Hit:Controller.Switch(64)=1:End Sub
Sub SW64_unHit:Controller.Switch(64)=0:End Sub
Sub SW74_Hit:Controller.Switch(74)=1:End Sub
Sub SW74_unHit:Controller.Switch(74)=0:End Sub

Sub SW65_Hit:Controller.Switch(65)=1:End Sub
Sub SW65_unHit:Controller.Switch(65)=0:End Sub
Sub SW75_Hit:Controller.Switch(75)=1:End Sub
Sub SW75_unHit:Controller.Switch(75)=0:End Sub

'Targets
Sub sw23_Hit:vpmTimer.PulseSw (23):End Sub
Sub sw33_Hit:vpmTimer.PulseSw (33):End Sub
Sub sw43_Hit:vpmTimer.PulseSw (43):End Sub

Sub sw52_Hit:vpmTimer.PulseSw (52):End Sub
Sub sw62_Hit:vpmTimer.PulseSw (62):End Sub
Sub sw72_Hit:vpmTimer.PulseSw (72):End Sub

Sub sw53_Hit:vpmTimer.PulseSw (53):End Sub
Sub sw63_Hit:vpmTimer.PulseSw (63):End Sub
Sub sw73_Hit:vpmTimer.PulseSw (73):End Sub

Sub sw02_Hit:vpmTimer.PulseSw (02):End Sub
Sub sw03_Hit:vpmTimer.PulseSw (03):End Sub

'Drop Targets
Sub SW40_Hit:vpmTimer.PulseSw (40):End Sub
Sub SW50_Hit:vpmTimer.PulseSw (50):End Sub
Sub SW60_Hit:vpmTimer.PulseSw (60):End Sub
Sub SW70_Hit:vpmTimer.PulseSw (70):End Sub

'Rollunder Gates

Sub SW34_Hit:vpmTimer.PulseSw (34):End Sub

'Spinners

Sub Sw24_Spin
  vpmTimer.PulseSw (24)
  PlaySound "fx_spinner", 0, .25, AudioPan(sw53), 0.25, 0, 0, 1, AudioFade(sw53)
End Sub


Sub SolKnocker(Enabled)
    If Enabled Then PlaySoundAt SoundFX("Knock",DOFKnocker), Plunger
End Sub




'**********Rubber Animations

sub sw55a_hit
    vpmtimer.PulseSw(55)
  Rsw55a.visible=0
  Rsw55a1.visible=1
  me.uservalue=1
  Me.timerenabled=1
end sub

sub sw55a_timer                 'default 50 timer
  select case me.uservalue
    Case 1: Rsw55A1.visible=0: Rsw55A.visible=1
    case 2: Rsw55A.visible=0: Rsw55A2.visible=1
    Case 3: Rsw55A2.visible=0: Rsw55A.visible=1: Me.timerenabled=0
  end Select
  me.uservalue=me.uservalue+1
end sub

sub sw55b_hit
    vpmtimer.PulseSw(55)
  Rsw55b.visible=0
  Rsw55b1.visible=1
  me.uservalue=1
  Me.timerenabled=1
end sub

sub sw55b_timer                 'default 50 timer
  select case me.uservalue
    Case 1: Rsw55b1.visible=0: Rsw55b.visible=1
    case 2: Rsw55b.visible=0: Rsw55b2.visible=1
    Case 3: Rsw55b2.visible=0: Rsw55b.visible=1: Me.timerenabled=0
  end Select
  me.uservalue=me.uservalue+1
end sub

'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep, Tstep, URstep


Sub RightSlingShot_Slingshot
    PlaySoundAt SoundFX("fx_slingshot",DOFContactors), slingR
    DOF 105, DOFPulse
    vpmtimer.PulseSw(25)
    RSling.Visible = 0
    RSling1.Visible = 1
    slingR.objroty = -15
    RStep = 0
    RightSlingShot.TimerEnabled = 1
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
    DOF 104, DOFPulse
    vpmtimer.pulsesw(25)
    LSling.Visible = 0
    LSling1.Visible = 1
    slingL.objroty = 15
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:slingL.objroty = 7
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:slingL.objroty=0:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

' *******************************************************************************************************
' Positional Sound Playback Functions by DJRobX and Rothbauerw
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

Sub PlaySoundAtBOTBallZ(sound, BOT)
    PlaySound sound, 0, ABS(BOT.velz)/17, Pan(BOT), 0, Pitch(BOT), 1, 0, AudioFade(BOT)
End Sub

' play a looping sound at a location with volume
Sub PlayLoopSoundAtVol(sound, tableobj, Vol)
  PlaySound sound, -1, Vol, AudioPan(tableobj), 0, 0, 1, 0, AudioFade(tableobj)
End Sub

'*********************************************************************
'                     Supporting Ball & Sound Functions
'*********************************************************************

Function RndNum(min, max)
    RndNum = Int(Rnd() * (max-min + 1) ) + min ' Sets a random number between min and max
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
  Vol = Csng(BallVel(ball) ^2 / VolDiv)
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
'      Ball Rolling Sounds by JP
'*****************************************

Const tnob = 15 ' total number of balls
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
    BallShadow(b).X = BOT(b).X-maxXoffset*(1-(Bot(b).X)/(Arena.Width/2))
    BallShadow(b).Y = BOT(b).Y + 10
    If BOT(b).Z > 0 and BOT(b).Z < 30 Then
      BallShadow(b).visible = 1
    Else
      BallShadow(b).visible = 0
    End If
  Next
End Sub

'Gottlieb Arena
'added by Inkochnito
Sub editDips
  Dim vpmDips : Set vpmDips = New cvpmDips
  With vpmDips
    .AddForm  700,400,"Arena - DIP switches"
    .AddFrame 2,4,190,"Maximum credits",49152,Array("8 credits",0,"10 credits",32768,"15 credits",&H00004000,"20 credits",49152)'dip 15&16
    .AddFrame 2,80,190,"Coin chute 1 and 2 control",&H00002000,Array("seperate",0,"same",&H00002000)'dip 14
    .AddFrame 2,126,190,"Playfield special",&H00200000,Array("replay",0,"extra ball",&H00200000)'dip 22
    .AddFrame 2,172,190,"High games to date control",&H00000020,Array("no effect",0,"reset high games 2-5 on power off",&H00000020)'dip 6
    .AddFrame 2,218,190,"Auto-percentage control",&H00000080,Array("disabled (normal high score mode)",0,"enabled",&H00000080)'dip 8
    .AddFrame 2,264,190,"Top 'GUARDS' targets",&H40000000,Array("are reset every ball",0,"are memorized ball to ball",&H40000000)'dip 31
    .AddFrame 2,310,190,"Completing all 6 spot targets",&H80000000,Array("2X lights special",0,"1X lights special",&H80000000)'dip 32
    .AddFrame 205,4,190,"High game to date awards",&H00C00000,Array("not displayed and no award",0,"displayed and no award",&H00800000,"displayed and 2 replays",&H00400000,"displayed and 3 replays",&H00C00000)'dip 23&24
    .AddFrame 205,80,190,"Balls per game",&H01000000,Array("5 balls",0,"3 balls",&H01000000)'dip 25
    .AddFrame 205,126,190,"Replay limit",&H04000000,Array("no limit",0,"one per game",&H04000000)'dip 27
    .AddFrame 205,172,190,"Novelty",&H08000000,Array("normal",0,"extra ball and replay scores 500K",&H08000000)'dip 28
    .AddFrame 205,218,190,"Game mode",&H10000000,Array("replay",0,"extra ball",&H10000000)'dip 29
    .AddFrame 205,264,190,"3rd coin chute credits control",&H20000000,Array("no effect",0,"add 9",&H20000000)'dip 30
    .AddChk 205,316,180,Array("Match feature",&H02000000)'dip 26
    .AddChk 205,331,190,Array("Attract sound",&H00000040)'dip 7
    .AddLabel 50,360,300,20,"After hitting OK, press F3 to reset game with new settings."
    .ViewDips
  End With
End Sub
Set vpmShowDips = GetRef("editDips")

'Finding an individual dip state based on scapino's Strikes and spares dip code - from unclewillys pinball pool


 Sub DipsTimer_Timer()
  Dim TheDips(32)
  Dim BPG, hsaward1, hsaward2, ebplay
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
  DipsNumber = Controller.Dip(2)
  TheDips(24) = Int(DipsNumber/128)
  If TheDips(24) = 1 then DipsNumber = DipsNumber - 128 end if
  TheDips(23) = Int(DipsNumber/64)
  If TheDips(23) = 1 then DipsNumber = DipsNumber - 64 end if
  TheDips(22) = Int(DipsNumber/32)
  If TheDips(22) = 1 then DipsNumber = DipsNumber - 32 end if
  TheDips(21) = Int(DipsNumber/16)
  If TheDips(21) = 1 then DipsNumber = DipsNumber - 16 end if
  TheDips(20) = Int(DipsNumber/8)
  If TheDips(20) = 1 then DipsNumber = DipsNumber - 8 end if
  TheDips(19) = Int(DipsNumber/4)
  If TheDips(19) = 1 then DipsNumber = DipsNumber - 4 end if
  TheDips(18) = Int(DipsNumber/2)
  If TheDips(18) = 1 then DipsNumber = DipsNumber - 2 end if
  TheDips(17) = Int(DipsNumber)

  hsaward1 = TheDips(23)
  hsaward2 = TheDips(24)




  BPG = TheDips(25)
  If BPG = 1 then
    instcard.image="InstCard3Balls"
    Else
    instcard.image="InstCard5Balls"
  End if
  if hsaward1 = 1 then
    repcard.image="replaycard1"
    Else
    repcard.image="replaycard2"
  end If
' DipsTimer.enabled=0
 End Sub

Sub Arena_Paused:Controller.Pause = 1:End Sub
Sub Arena_unPaused:Controller.Pause = 0:End Sub

Sub Arena_Exit
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

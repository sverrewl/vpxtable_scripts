Option Explicit
Randomize
' Thalamus - added more randomness to kickers
Const cGameName = "alienstr"

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "Can't open controller.vbs"
On Error Goto 0

LoadVPM "01210000", "sys80.VBS", 3.1


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

Dim bsTrough, bslLock, objekt, xx
Dim showdmd


Sub AlienStr_Init
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
     .Hidden = 1
         On Error Resume Next
         .Run GetPlayerHWnd
         If Err Then MsgBox Err.Description
         On Error Goto 0
     End With

  If B2SOn then
    for each objekt in backdropstuff : objekt.visible = 0 : next
  End If

  Intensity 'sets GI brightness depending on day/night slider settings

' Thalamus : Was missing 'vpminit me'
  vpminit me

    PinMAMETimer.Interval = PinMAMEInterval
    PinMAMETimer.Enabled = 1

'************************Trough

Kloadgame.uservalue=0
Kloadgame.timerenabled=1

' sw14.CreateBall
' Controller.Switch(14) = 1
''  sw44.CreateBall
''  Controller.Switch(44) = 1

'Kickers

    Set bslLock=New cvpmBallStack
    with bslLock
        .InitSaucer sw04,04,-164,8.5
        .InitExitSnd Soundfx("holekick",DOFContactors), Soundfx("HoleKick",DOFContactors)
        .KickForceVar = 9
        .KickAngleVar = 14
    end with


' Nudging
  vpmNudge.TiltSwitch = 57
    vpmNudge.Sensitivity = 1
    vpmNudge.TiltObj = Array(leftslingshot, rightslingshot, bumper1, bumper2, bumper3)


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

End Sub

sub Kloadgame_timer
  kloadgame.createball
  Kloadgame.kick 0,0
  kloadgame.uservalue=kloadgame.uservalue+1
  if kloadgame.uservalue=2 then me.timerenabled=0
end sub


'************************************************
' Solenoids
'************************************************
SolCallback(1) =    "SolLLock"


SolCallback(8) =    "solknocker"
SolCallback(9) =    "solballrelease"

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
     If Enabled Then
        PlaySound SoundFX("fx_Flipperup",DOFFlippers):LeftFlipper.RotateToEnd
    controller.Switch(48)=1
     Else
        PlaySound SoundFX("fx_Flipperdown",DOFFlippers):LeftFlipper.RotateToStart
    controller.Switch(48)=0
     End If
  End Sub


Sub SolRFlipper(Enabled)
     If Enabled Then
         PlaySound SoundFX("fx_Flipperup",DOFFlippers):RightFlipper.RotateToEnd
    controller.Switch(46)=1
     Else
         PlaySound SoundFX("fx_Flipperdown",DOFFlippers):RightFlipper.RotateToStart
    controller.Switch(46)=0
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



Sub FlipperTimer_Timer
  Dim PI: PI=3.1415926
  FreflectS.visible = L29.state
  FreflectT.visible = L30.state
  FreflectA.visible = L31.state
  FreflectR.visible = L32.state

  LFlip.RotY = LeftFlipper.CurrentAngle
  RFlip.RotY = RightFlipper.CurrentAngle
  Pgate.rotx = -Gate.currentangle*0.5
  Pgate1.rotx = -Gate1.currentangle*0.5

  Dim SpinnerRadius: SpinnerRadius=7

  SpinnerRod.TransZ = (cos((sw53.CurrentAngle + 180) * (PI/180))+1) * SpinnerRadius
  SpinnerRod.TransY = sin((sw53.CurrentAngle) * (PI/180)) * -SpinnerRadius

'testbox.text = sw53.currentangle
'testbox1.text = "transZ "&SpinnerRod.TransZ
'testbox2.text = "transY "&SpinnerRod.TransY

  if FlipperShadows=1 then
    FlipperLSh.RotZ = LeftFlipper.currentangle
    FlipperRSh.RotZ = RightFlipper.currentangle
  end if
End Sub


' Ball locks / kickers

Sub sw04_Hit:PlaySound "holein":bslLock.AddBall 0:End Sub


Sub SolLLock(enabled)
  If enabled Then
    bslLock.ExitSol_On
    DOF 106, DOFpulse
    LeftKickTimer.uservalue = 0
    Pkickarm.RotZ = 15
    LeftKickTimer.Enabled = 1
  End If
End Sub

Sub LeftKickTimer_timer
  select case me.uservalue
    case 5:
    Pkickarm.rotz=0
    me.enabled=0
  end Select
  me.uservalue = me.uservalue+1
End Sub



'******************************************************
'     DRAIN & TROUGH BASED ON NFOZZY'S via bord's pink panther
'******************************************************

Sub sw14_Hit():Controller.Switch(14) = 1:End Sub
Sub sw14_UnHit():Controller.Switch(14) = 0:End Sub

Sub sw44_Hit(): PlaySoundat "Switch", sw44: Controller.Switch(44) = 1: End Sub
Sub sw44_UnHit(): Controller.Switch(44) = 0: End Sub

Sub solballrelease(enabled)
  If enabled Then
    PlaySoundat SoundFX("holekick",DOFContactors), sw44
    sw44.kick 60, 40
  End If
End Sub

Set LampCallback = GetRef("UpdateMultipleLamps")
Sub UpdateMultipleLamps
    if controller.lamp(12)=true and sw14.BallCntOver > 0 then
    sw14.kick 60, 3
    PlaySoundAt SoundFX("ballrelease",DOFContactors), sw14
    end if

    If Controller.Lamp(11) Then   'Game Over triggers match and BIP
        GOBox.text="GAME OVER"
      Else
        GOBox.text=""
    End If


    If Controller.Lamp(1)  Then         'Tilt
        TILTBox.text="TILT"
      Else
        TILTBox.text=""
    End If


    If Controller.Lamp(10) Then       'HIGH SCORE TO DATE
        HStoDateBox.text="HIGH SCORE TO DATE"
      Else
        HStoDateBox.text=""
    End If

End Sub

Dim Digits(32)

'Score displays

'Digits(0)=Array(a1,a2,a3,a4,a5,a6,a7,n,a8)
'Digits(1)=Array(a9,a10,a11,a12,a13,a14,a15,n,a16)
'Digits(2)=Array(a17,a18,a19,a20,a21,a22,a23,n,a24)
'Digits(3)=Array(a25,a26,a27,a28,a29,a30,a31,n,a32)
'Digits(4)=Array(a33,a34,a35,a36,a37,a38,a39,n,a40)
'Digits(5)=Array(a41,a42,a43,a44,a45,a46,a47,n,a48)
'Digits(6)=Array(a49,a50,a51,a52,a53,a54,a55,n,a56)
'Digits(7)=Array(a57,a58,a59,a60,a61,a62,a63,n,a64)
'Digits(8)=Array(a65,a66,a67,a68,a69,a70,a71,n,a72)
'Digits(9)=Array(a73,a74,a75,a76,a77,a78,a79,n,a80)
'Digits(10)=Array(a81,a82,a83,a84,a85,a86,a87,n,a88)
'Digits(11)=Array(a89,a90,a91,a92,a93,a94,a95,n,a96)
'Digits(12)=Array(a97,a98,a99,a100,a101,a102,a103,n,a104)
'Digits(13)=Array(a105,a106,a107,a108,a109,a110,a111,n,a112)
'Digits(14)=Array(a113,a114,a115,a116,a117,a118,a119,n,a120)
'Digits(15)=Array(a121,a122,a123,a124,a125,a126,a127,n,a128)
'Digits(16)=Array(a129,a130,a131,a132,a133,a134,a135,n,a136)
'Digits(17)=Array(a137,a138,a139,a140,a141,a142,a143,n,a144)
'Digits(18)=Array(a145,a146,a147,a148,a149,a150,a151,n,a152)
'Digits(19)=Array(a153,a154,a155,a156,a157,a158,a159,n,a160)
'Digits(20)=Array(a161,a162,a163,a164,a165,a166,a167,n,a168)
'Digits(21)=Array(a169,a170,a171,a172,a173,a174,a175,n,a176)
'Digits(22)=Array(a177,a178,a179,a180,a181,a182,a183,n,a184)
'Digits(23)=Array(a185,a186,a187,a188,a189,a190,a191,n,a192)
'Digits(24)=Array(a193,a194,a195,a196,a197,a198,a199,n,a200)
'Digits(25)=Array(a201,a202,a203,a204,a205,a206,a207,n,a208)
'Digits(26)=Array(a209,a210,a211,a212,a213,a214,a215,n,a216)
'Digits(27)=Array(a217,a218,a219,a220,a221,a222,a223,n,a224)
'
''Ball in Play and Credit displays
'
'Digits(30)=Array(e00,e01,e02,e03,e04,e05,e06,n)
'Digits(31)=Array(e10,e11,e12,e13,e14,e15,e16,n)
'Digits(28)=Array(f00,f01,f02,f03,f04,f05,f06,n)
'Digits(29)=Array(f10,f11,f12,f13,f14,f15,f16,n)


Sub DisplayTimer_Timer
  Exit Sub
End Sub

'Sub DisplayTimer_Timer
'    Dim ChgLED,ii,num,chg,stat, obj
'    ChgLed = Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
'    If Not IsEmpty(ChgLED) Then
''        If not b2son Then
'        For ii = 0 To UBound(chgLED)
'            num = chgLED(ii, 0) : chg = chgLED(ii, 1) : stat = chgLED(ii, 2)
'            if (num < 32 ) then
'                For Each obj In Digits(num)
'                    If chg And 1 Then obj.State = stat And 1
'                    chg = chg\2 : stat = stat\2
'                Next
'            else
'
'            end if
'        next
''        end if
'    end if
'End Sub




'**********************************************************************************************************
'Plunger code
'**********************************************************************************************************

Sub AlienStr_KeyDown(ByVal KeyCode)

    If keycode = PlungerKey Then Plunger.Pullback:PlaySoundAt "plungerpull", Plunger


    If KeyDownHandler(keycode) Then Exit Sub

End Sub

Sub AlienStr_KeyUp(ByVal KeyCode)

    If keycode = PlungerKey Then
    Plunger.Fire
    if BallHome.BallCntOver>0 then
      PlaySoundAt "plungerreleaseball", Plunger
      else
      PlaySoundAt "plungerreleasefree", Plunger
    end if
  end if



    If KeyUpHandler(keycode) Then Exit Sub

End Sub



'Bumpers

Sub bumper1_Hit : vpmTimer.PulseSw 45 : PlaySoundAt SoundFX("fx_bumper4",DOFContactors), Bumper1: DOF 102, DOFPulse:End Sub
Sub bumper2_Hit : vpmTimer.PulseSw 45 : PlaySoundAt SoundFX("fx_bumper4",DOFContactors), Bumper2: DOF 103, DOFPulse:End Sub
Sub bumper3_Hit : vpmTimer.PulseSw 45 : PlaySoundAt SoundFX("fx_bumper4",DOFContactors), Bumper3: DOF 101, DOFPulse:End Sub


'Wire Triggers
Sub SW00_Hit:Controller.Switch(00)=1::End Sub
Sub SW00_unHit:Controller.Switch(00)=0:End Sub
Sub SW10_Hit:Controller.Switch(10)=1::End Sub
Sub SW10_unHit:Controller.Switch(10)=0:End Sub
Sub SW40_Hit:Controller.Switch(40)=1::End Sub
Sub SW40_unHit:Controller.Switch(40)=0:End Sub
Sub SW50_Hit:Controller.Switch(50)=1::End Sub
Sub SW50_unHit:Controller.Switch(50)=0:End Sub
Sub SW13_Hit:Controller.Switch(13)=1:End Sub
Sub SW13_unHit:Controller.Switch(13)=0:End Sub
Sub SW43_Hit:Controller.Switch(43)=1:End Sub
Sub SW43_unHit:Controller.Switch(43)=0:End Sub
Sub SW42_Hit:Controller.Switch(42)=1:End Sub
Sub SW42_unHit:Controller.Switch(42)=0:End Sub
' Thalamus : This sub is used twice - this means ... this one IS NOT USED
' Not a issue though, they are the same - missed another switch ?
' Sub SW12_Hit:Controller.Switch(12)=1:End Sub
' Sub SW12_unHit:Controller.Switch(12)=0:End Sub
Sub SW12_Hit:Controller.Switch(12)=1:End Sub
Sub SW12_unHit:Controller.Switch(12)=0:End Sub


'Targets
Sub sw01_Hit:vpmTimer.PulseSw (01):End Sub
Sub sw11_Hit:vpmTimer.PulseSw (11):End Sub
Sub sw41_Hit:vpmTimer.PulseSw (41):End Sub
Sub sw51_Hit:vpmTimer.PulseSw (51):End Sub
Sub sw02_Hit:vpmTimer.PulseSw (02):End Sub
Sub sw03_Hit:vpmTimer.PulseSw (03):End Sub

'Spinners

Sub Sw53_Spin
  vpmTimer.PulseSw (53)
  PlaySound "fx_spinner", 0, .25, AudioPan(sw53), 0.25, 0, 0, 1, AudioFade(sw53)
End Sub

'Gate Switch

Sub sw52_hit
  vpmTimer.PulseSw (52)
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
    vpmtimer.PulseSw(55)
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
    vpmtimer.pulsesw(55)
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
    tmp = tableobj.y * 2 / AlienStr.height-1
    If tmp > 0 Then
    AudioFade = Csng(tmp ^10)
    Else
        AudioFade = Csng(-((- tmp) ^10) )
    End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "table1" is the name of the table
    Dim tmp
    tmp = tableobj.x * 2 / AlienStr.width-1
    If tmp > 0 Then
        AudioPan = Csng(tmp ^10)
    Else
        AudioPan = Csng(-((- tmp) ^10) )
    End If
End Function

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
    Vol = Csng(BallVel(ball) ^2 / 1000)
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
'   testbox.text = "gitcatcher.state = " & input    'debug
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
    BallShadow(b).X = BOT(b).X-maxXoffset*(1-(Bot(b).X)/(AlienStr.Width/2))
    BallShadow(b).Y = BOT(b).Y + 10
    If BOT(b).Z > 0 and BOT(b).Z < 30 Then
      BallShadow(b).visible = 1
    Else
      BallShadow(b).visible = 0
    End If
  Next
End Sub

'*****************************************
'Gottlieb Alien Star dip Switch
' by Inkochnito
'*****************************************
Sub editDips
  Dim vpmDips : Set vpmDips = New cvpmDips
  With vpmDips
    .AddForm 700,400,"Alien Star - DIP switches"
    .AddFrame 2,2,190,"Maximum credits",49152,Array("8 credits",0,"10 credits",32768,"15 credits",&H00004000,"20 credits",49152)'dip 15&16
    .AddFrame 2,80,190,"Coin chute 1 and 2 control",&H00002000,Array("seperate",0,"same",&H00002000)'dip 14
    .AddFrame 2,126,190,"Playfield special",&H00200000,Array("replay",0,"extra ball",&H00200000)'dip 22
    .AddFrame 2,172,190,"Game mode",&H10000000,Array("replay",0,"extra ball",&H10000000)'dip 29
    .AddChk 2,227,190,Array("Match feature",&H02000000)'dip 26
    .AddChk 2,248,190,Array("Background sound",&H40000000)'dip 31
    .AddFrame 205,2,190,"High score to date awards",&H00C00000,Array("not displayed and no award",0,"displayed and no award",&H00800000,"displayed and 2 replays",&H00400000,"displayed and 3 replays",&H00C00000)'dip 23&24
    .AddFrame 205,80,190,"Balls per game",&H01000000,Array("5 balls",0,"3 balls",&H01000000)'dip 25
    .AddFrame 205,126,190,"Replay limit",&H04000000,Array("no limit",0,"one per ball",&H04000000)'dip 27
    .AddFrame 205,172,190,"Novelty",&H08000000,Array("normal",0,"points",&H0800000)'dip 28
    .AddFrame 205,218,190,"3rd coin chute credits control",&H20000000,Array("no effect",0,"add 9",&H20000000)'dip 30
    .AddLabel 50,270,300,20,"After hitting OK, press F3 to reset game with new settings."
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
  ebplay = TheDips(29)

'testbox.text = hsaward1
'testbox1.text = hsaward2
'testbox2.text = ebplay

  BPG = TheDips(25)
  If BPG = 1 then
    instcard.image="InstCard3Balls"
    Else
    instcard.image="InstCard5Balls"
  End if
  if ebplay = 0 then
    if hsaward1 = 1 Then
      if hsaward2 = 1 then
        repcard.image="replaycard1"
        Else
        repcard.image="replaycard2"
      end If
      Else
      repcard.image="replaycard0"
    end if
    else
    repcard.image="replaycard0"
  end if
' DipsTimer.enabled=0
 End Sub

Sub AlienStr_Paused:Controller.Pause = 1:End Sub
Sub AlienStr_unPaused:Controller.Pause = 0:End Sub

Sub AlienStr_Exit
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

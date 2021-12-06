'****Conet Williams(1985)
'****Original By Unclewilly w/Textures By Grizz
'VPX version by Sliderpoint w/main ramp textured by Flupper
Option Explicit
Randomize

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "Can't open controller.vbs"
On Error Goto 0

'Options
Dim FlipperShadows, BallShadows

FlipperShadows = 1 ' change to 0 to turn off flipper shadows
BallShadows = 1 ' change to 0 to turn off ball shadow

'Variables
Dim xx, cGameName
Const UseSolenoids = 2
Const UseLamps = 1
Const UseGI = 0
Const UseSync = 0
Const HandleMech = 0
Const SSolenoidOn = "Solenoid"
Const SSolenoidOff = ""
Const SCoin = "CoinIn"
Dim DesktopMode: DesktopMode = Table1.ShowDT
cGameName = "comet_l5"

LoadVPM "01560000", "S11.VBS", 3.26

'Table Init
  Sub Table1_Init
    vpmInit Me
          With Controller
                .GameName = cGameName
                .SplashInfoLine = "Comet (Williams 1985)" & vbNewLine & "for VPX by Sliderpoint"
                .HandleMechanics = 0
                .HandleKeyboard = 0
                .ShowDMDOnly = 1
                .ShowFrame = 0
                .ShowTitle = 0
        End With
          Controller.Hidden = 0
        On Error Resume Next

           Controller.Run
           If Err Then MsgBox Err.Description
             On Error Goto 0
 'Nudging
    vpmNudge.TiltSwitch=swTilt
    vpmNudge.Sensitivity=1
    vpmNudge.TiltObj=Array(Bumper1,Bumper2,Bumper3,LeftSlingshot,RightSlingshot)

  Outhole.CreateSizedBallWithMass 25, 1.65
  Outhole.kick 0,0

  vpmMapLights AllLamps
  Intensity 'sets GI brightness based on day/night slider
  If BallShadows = 0 Then
    BallShadowUpdate.Enabled = 0
  End If
  If DesktopMode = False Then
    RailLeft.visible = 0
    RailRight.visible = 0
'   SideWalls.visible = 0
  End if
  End Sub

Sub Table1_Paused:Controller.Pause = 1:End Sub
Sub Table1_unPaused:Controller.Pause = 0:End Sub
sub Table1_Exit:Controller.Stop:end sub

Sub Table1_KeyDown(ByVal keycode)
  If keycode = PlungerKey Then Plunger.PullBack
    If keycode = LeftTiltKey Then playSound "nudge_left"
    If keycode = RightTiltKey Then playSound "nudge_right"
    If keycode = CenterTiltKey Then PlaySound "nudge_forward"
    If vpmKeyDown(keycode) Then Exit Sub
End Sub

Sub Table1_KeyUp(ByVal keycode)
  If keycode = PlungerKey Then Plunger.Fire:PlaySound "plunger"
  If vpmKeyUp(keycode) Then Exit Sub
End Sub

'Solenoids
    SolCallback(1) = "Outholekick"
    SolCallback(2) = "DTReset"
    SolCallback(3) = "RabbitKick"
    SolCallback(4) = "CSFlash"    'Corkscrew Flash
  SolCallback(5) = "CycleFlash" 'Cycle Flash
    SolCallback(6) = "CycleJumpKick"
' SolCallback(7) = "" 'Player3 Flasher
' SolCallback(8) = "" 'Player1 Flasher
' SolCallback(9) = "" 'Player4 Flasher
' SolCallback(10) = ""'Player2 Flasher
    SolCallback(11) = "solGI"
    SolCallback(15) = "vpmSolSound ""Knocker"","

Sub OutholeKick(enabled)
  If enabled Then
  Outhole.kick 65, 20
  Playsound SoundFX("BallRelease",DOFContactors), 0, 1, AudioPan(Outhole), 0,0,0, 1, AudioFade(Outhole)
  End If
end Sub

Sub Outhole_Hit:PlaySound "Scoopenter", 0, 1, AudioPan(Outhole), 0,0,0, 1, AudioFade(Outhole):controller.switch(45) = 1:End Sub
Sub Outhole_unHit:Controller.Switch(45) = 0: End Sub

Sub RabbitKick(enabled)
  If enabled Then
  Sw24.kick 175, 10
  Playsound SoundFX("Solenoid",DOFContactors), 0, 1, AudioPan(sw24), 0,0,0, 1, AudioFade(sw24)
  End If
end Sub

Sub sw24_Hit:PlaySound "kicker_enter", 0, 1, AudioPan(sw24), 0,0,0, 1, AudioFade(sw24):controller.switch(24) = 1:End Sub
Sub sw24_unHit:Controller.Switch(24) = 0: End Sub

Sub CSFlash(enabled)
  If Enabled Then
    CorkScrewFlasher.state = 1
    CorkScrewFlasher1.state = 1
    CorkScrewFlasher2.state = 1
  Else
    CorkScrewFlasher.State = 0
    CorkScrewFlasher1.state = 0
    CorkScrewFlasher2.state = 0
  End If
End Sub

Sub CycleFlash(enabled)
  If Enabled Then
    CycleFlasher1.state = 1
    CycleFlasher2.state = 1
    CycleFlasher3.state = 1
    CycleFlasher4.visible = 1
  Else
    CycleFlasher1.State = 0
    CycleFlasher2.State = 0
    CycleFlasher3.state = 0
    CycleFlasher4.visible = 0
  End If
End Sub

Sub CycleJumpKick(enabled)
  If enabled Then
  sw25.kick 260, 18
  Playsound SoundFX("Solenoid",DOFContactors), 0, 1, AudioPan(sw25), 0,0,0, 1, AudioFade(sw25)
  End If
end Sub

Sub sw25_Hit:PlaySound "kicker_enter", 0, 1, AudioPan(sw25), 0,0,0, 1, AudioFade(sw25):controller.switch(25) = 1:End Sub
Sub sw25_unHit:Controller.Switch(25) = 0: End Sub

Sub DTReset (enabled)
  sw29.isdropped = 0
  controller.switch(29) = 0
  Playsound SoundFX("Solenoid",DOFContactors), 0, 1, AudioPan(sw29), 0,0,0, 1, AudioFade(sw29)
End Sub

'Flippers
SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
    If Enabled Then
    PlaySound SoundFX("fx_FlipperUp",DOFFlippers), 0, 1, AudioPan(LeftFlipper), 0,0,0,1,AudioFade(LeftFlipper)
    LeftFlipper.RotateToEnd
    Else
    PlaySound SoundFX("fx_FlipperDown",DOFFlippers), 0, 1, AudioPan(LeftFlipper), 0,0,0,1,AudioFade(LeftFlipper)
    LeftFlipper.RotateToStart
    End If
End Sub

Sub SolRFlipper(Enabled)
    If Enabled Then
    PlaySound SoundFX("fx_FlipperUp",DOFFlippers), 0, 1, AudioPan(RightFlipper), 0,0,0,1,AudioFade(RightFlipper)
    RightFlipper.RotateToEnd
    Else
    PlaySound SoundFX("fx_FlipperDown",DOFFlippers), 0, 1, AudioPan(RightFlipper), 0,0,0,1,AudioFade(RightFlipper)
    RightFlipper.RotateToStart
    End If
End Sub

Sub LeftFlipper_Collide(parm):RandomFlipperSound:End Sub
Sub RightFlipper_Collide(parm):RandomFlipperSound:End Sub

'Sling Shots
Dim RStep, Lstep

Sub RightSlingShot_Slingshot
    PlaySound SoundFX("rSling",DOFContactors), 0,1, 0.05,0.05
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.TransZ = -20
    RStep = 0
    RightSlingShot.TimerEnabled = 1
  vpmTimer.PulseSw 48
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.TransZ = -10
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.TransZ = 0:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
    PlaySound SoundFX("Lsling",DOFContactors), 0,1, -0.05,0.05
    LSling.Visible = 0
    LSling1.Visible = 1
    sling2.TransZ = -20
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
  vpmTimer.PulseSw 47
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.TransZ = -10
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.TransZ = 0:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

'GI
 Sub solGI(enabled)
    If enabled then
     For each xx in GIBulbs:xx.State=0:Next
     For each xx in GIBulbs2:xx.State=0:Next
     SideWalls.image = "SideWallsEarlySSMap"
    Playsound "Flasher_relay_off"
    else
       For each xx in GIBulbs:xx.State=1:Next
       For each xx in GIBulbs2:xx.State=1:Next
     SideWalls.image = "SideWallsEarlySSOn"
     Playsound "Flasher_relay_on"
    end if
End Sub

Dim GILevel, DayNight
DayNight = Table1.NightDay
Sub Intensity
  If DayNight > 10 Then
  GILevel = (Round((1/DayNight),2))*12
  Else
  GILevel = 1
  End If

  For each xx in GIbulbs: xx.IntensityScale = xx.IntensityScale * (GILevel): Next
  For each xx in GIbulbs2: xx.IntensityScale = xx.IntensityScale * (GILevel): Next
End Sub
'End GI

  Sub sw33_Hit():RandomRubberSound:vpmTimer.PulseSw 33:End Sub
  Sub sw34_Hit():RandomRubberSound:vpmTimer.PulseSw 34:End Sub
  Sub sw35_Hit():RandomRubberSound:vpmTimer.PulseSw 35:End Sub
  Sub sw36_Hit():RandomRubberSound:vpmTimer.PulseSw 36:End Sub
  Sub sw37_Hit():RandomRubberSound:vpmTimer.PulseSw 37:End Sub
  Sub sw38_Hit():RandomRubberSound:vpmTimer.PulseSw 38:End Sub
  Sub sw39_Hit():RandomRubberSound:vpmTimer.PulseSw 39:End Sub

'Bumpers
   Sub Bumper1_Hit:vpmTimer.PulseSw 40:PlaySound SoundFX("BumperC",DOFContactors), 0,1,AudioPan(Bumper1),0,0,0,1,AudioFade(Bumper1):End Sub
   Sub Bumper2_Hit:vpmTimer.PulseSw 41:PlaySound SoundFX("BumperC",DOFContactors), 0,1,AudioPan(Bumper2),0,0,0,1,AudioFade(Bumper2):End Sub
   Sub Bumper3_Hit:vpmTimer.PulseSw 42:PlaySound SoundFX("BumperC",DOFContactors), 0,1,AudioPan(Bumper3),0,0,0,1,AudioFade(Bumper3):End Sub

'Rollover & Ramp Switches
    Sub sw9_Hit:Controller.Switch(9) = 1:PlaySound "rollover", 0, Vol(ActiveBall)*10, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall):End Sub
    Sub sw9_UnHit:Controller.Switch(9) = 0:End Sub
    Sub sw19_Hit:vpmTimer.PulseSw 19:PlaySound "Gate", 0, Vol(ActiveBall)*10, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall):End Sub
    Sub sw20_Hit:Controller.Switch(20) = 1:PlaySound "rollover", 0, Vol(ActiveBall)*10, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall):End Sub
    Sub sw20_UnHit:Controller.Switch(20) = 0:End Sub
    Sub sw21_Hit:Controller.Switch(21) = 1:PlaySound "rollover", 0, Vol(ActiveBall)*10, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall):End Sub
    Sub sw21_UnHit:Controller.Switch(21) = 0:End Sub
    Sub sw22_Hit:Controller.Switch(22) = 1:PlaySound "rollover", 0, Vol(ActiveBall)*10, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall):End Sub
    Sub sw22_UnHit:Controller.Switch(22) = 0:End Sub
    Sub sw23_Hit:Controller.Switch(23) = 1:PlaySound "rollover", 0, Vol(ActiveBall)*10, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall):End Sub
    Sub sw23_UnHit:Controller.Switch(23) = 0:End Sub
    Sub sw26_Hit:Controller.Switch(26) = 1:PlaySound "rollover", 0, Vol(ActiveBall)*10, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall):End Sub
    Sub sw26_UnHit:Controller.Switch(26) = 0:End Sub
    Sub sw27_Hit:Controller.Switch(27) = 1:PlaySound "rollover", 0, Vol(ActiveBall)*10, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall):End Sub
    Sub sw27_UnHit:Controller.Switch(27) = 0:End Sub
    Sub sw31_Hit:Controller.Switch(31) = 1:PlaySound "rollover", 0, Vol(ActiveBall)*10, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall):End Sub
    Sub sw31_UnHit:Controller.Switch(31) = 0:End Sub
    Sub sw32_Hit:Controller.Switch(32) = 1:PlaySound "rollover", 0, Vol(ActiveBall)*10, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall):End Sub
    Sub sw32_UnHit:Controller.Switch(32) = 0:End Sub
    Sub sw43_Hit:Controller.Switch(43) = 1:PlaySound "rollover", 0, Vol(ActiveBall)*10, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall):End Sub
    Sub sw43_UnHit:Controller.Switch(43) = 0:End Sub
    Sub sw44_Hit:Controller.Switch(44) = 1:PlaySound "rollover", 0, Vol(ActiveBall)*10, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall):End Sub
    Sub sw44_UnHit:Controller.Switch(44) = 0:End Sub
    Sub sw46_Hit:Controller.Switch(46) = 1:End Sub
    Sub sw46_UnHit:Controller.Switch(46) = 0:End Sub
    Sub sw50_Hit:Controller.Switch(50) = 1:sw50.TimerEnabled = 1:dim i:for i=1 to 4:StopSound "fx_rampbump" & i:next:NextOrbitHit = Timer + 1: End Sub
  Sub sw50_Timer:PlaySoundAt "DROP_RIGHT", sw50: sw50.TimerEnabled = 0: End Sub
    Sub sw28_Hit
    vpmTimer.PulseSw 28
    PlaySound"Gate", 0, Vol(ActiveBall)*10, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    PlaySound"RampEntrywGate", 0, Vol(ActiveBall)*10, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End Sub

    Sub cycledrop_Hit(IDX):PlaySound"DROP_RIGHT":End Sub

'StandUp Targets
    Sub sw10_Hit:vpmTimer.PulseSw 10:PlaySound "target", 0, Vol(ActiveBall)*10, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall):End Sub
    Sub sw11_Hit:vpmTimer.PulseSw 11:PlaySound "target", 0, Vol(ActiveBall)*10, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall):End Sub
    Sub sw12_Hit:vpmTimer.PulseSw 12:PlaySound "target", 0, Vol(ActiveBall)*10, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall):End Sub
    Sub sw13_Hit:vpmTimer.PulseSw 13:PlaySound "target", 0, Vol(ActiveBall)*10, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall):End Sub
    Sub sw14_Hit:vpmTimer.PulseSw 14:PlaySound "target", 0, Vol(ActiveBall)*10, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall):End Sub
    Sub sw15_Hit:vpmTimer.PulseSw 15:PlaySound "target", 0, Vol(ActiveBall)*10, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall):End Sub
    Sub sw16_Hit:vpmTimer.PulseSw 16:PlaySound "target", 0, Vol(ActiveBall)*10, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall):End Sub
    Sub sw17_Hit:vpmTimer.PulseSw 17:PlaySound "target", 0, Vol(ActiveBall)*10, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall):End Sub
    Sub sw18_Hit:vpmTimer.PulseSw 18:PlaySound "target", 0, Vol(ActiveBall)*10, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall):End Sub

'DropTargets
    Sub sw29_Hit:sw29.isdropped = 1:controller.Switch(29) = 1:Playsound SoundFX("droptarget",DOFContactors), 0, 1, AudioPan(sw29), 0,0,0, 1, AudioFade(sw29):End Sub

'Misc
dim PI:PI=3.1415926
Sub MiscTimer_Timer
  sw28p.RotX = sw28.currentAngle
  sw19p.RotX = sw19.currentAngle
  Gate5p.RotY = -Gate5.currentAngle
  Gate4p.Roty = gate4.currentAngle
  Gate2p.RotY = gate2.currentAngle
  FlipperR.RotZ = RightFlipper.currentAngle
  FlipperL.RotZ = LeftFlipper.CurrentAngle
  L57.visible = L57b.State
  If FlipperShadows = 1 Then
    FlipperShadowL.RotZ = LeftFlipper.currentAngle
    FlipperShadowR.RotZ = RightFlipper.currentAngle
  End If
  SpinnerRod.TransZ = 1.05*(sin( (sw28.CurrentAngle+180) * (2*PI/360)) * 8)
  SpinnerRod.TransY = 2.5*(sin( (sw28.CurrentAngle) * (2*PI/360)) * 8)
  SpinnerRod2.TransZ = -2.5*(sin( (sw50Gate.CurrentAngle+180) * (2*PI/360)) * 8)
  SpinnerRod2.TransY = 1.05*(sin( (sw50Gate.CurrentAngle) * (2*PI/360)) * 8)
  sw50Wire.RotX = -SW50Gate.currentAngle
End Sub

If FlipperShadows = 1 Then
  FlipperShadowL.visible = 1
  FlipperShadowR.visible = 1
  Else
  FlipperShadowL.visible = 0
  FlipperShadowR.visible = 0
End If

'*****************************************
'      JP's VP10 Rolling Sounds
'*****************************************
Const tnob = 1 ' total number of balls
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
      If BallVel(BOT(b) ) > 1 AND BOT(b).z < 30 Then
        rolling(b) = True
        PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b)), AudioPan(BOT(b)), 0, Pitch(BOT(b)), 1, 0, AudioFade(BOT(b))
      Else ' Ball on raised ramp
        PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) )*.35, AudioPan(BOT(b) ), 0, Pitch(BOT(b) )-10000, 1, 0, AudioFade(BOT(b) )
        End If
        Else
            If rolling(b) = True Then
                StopSound("fx_ballrolling" & b)
                rolling(b) = False
            End If
        End If
    Next
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
  Vol =  (Round(BallVel(ball)/10,2))
End Function

Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
    Pitch = BallVel(ball) * 20
End Function

Function BallVel(ball) 'Calculates the ball speed
    BallVel = INT(SQR((ball.VelX ^2) + (ball.VelY ^2) ) )
End Function

'**********************
' Ball Collision Sound
'**********************
Sub OnBallBallCollision(ball1, ball2, velocity)
  PlaySound("collide5"), 0, Csng(velocity) ^2 / 2000, AudioPan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub

'Extra Sounds
  Sub Rubbers_Hit(IDX):RandomRubberSound:End Sub
  Sub Gates_Hit(IDX):PlaySound "Gate", 0, Vol(ActiveBall)*10, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall):End Sub
  Sub MetalWalls_Hit(IDX):RandomMetalSound:End Sub
  Sub WireGuides_Hit(IDX):RandomMetalSound:End Sub
  Sub RubberPosts_Hit(IDX):PlaySound "rubber", 0, Vol(ActiveBall)*10, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall):End Sub
  Sub RampSound2_Hit:PlaySound"RampMidway", 0, Vol(ActiveBall)*10, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall): End Sub
  Sub RampSound3_Hit:PlaySound"RampEnd", 0, Vol(ActiveBall)*10, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall):End Sub
  Sub RampSound4_Hit:PlaySound"RampMetalHit", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall):End Sub
'RustyCardores/DJRobX bump sounds on ramps.
Dim NextOrbitHit:NextOrbitHit = 0

Sub PlasticRampBumps_Hit(idx)
  if BallVel(ActiveBall) > .3 and Timer > NextOrbitHit then
    RandomBump 2, Pitch(ActiveBall)
    NextOrbitHit = Timer + .1 + (Rnd * .2)
  end if
End Sub

Sub RandomBump(voladj, freq)
  dim BumpSnd:BumpSnd= "fx_rampbump" & CStr(Int(Rnd*7)+1)
    PlaySound BumpSnd, 0, Vol(ActiveBall)*voladj, AudioPan(ActiveBall), 0, freq, 0, 1, AudioFade(ActiveBall)
End Sub
'/
Sub RandomFlipperSound
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySoundAtBall "flip_hit_1"
    Case 2 : PlaySoundAtBall "flip_hit_2"
    Case 3 : PlaySoundAtBall "flip_hit_3"
  End Select
End Sub

Sub RandomRubberSound
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySoundAtBall "rubber_hit_1"
    Case 2 : PlaySoundAtBall "rubber_hit_2"
    Case 3 : PlaySoundAtBall "rubber_hit_3"
  End Select
End Sub

Sub RandomMetalSound
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySoundAtBall "metalhit_heavy"
    Case 2 : PlaySoundAtBall "metalhit_medium"
    Case 3 : PlaySoundAtBall "metalhit_thin"
  End Select
End Sub

'*****************************************
' ninuzzu's BALL SHADOW
'*****************************************
Dim BallShadow
BallShadow = Array (BallShadow1)

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


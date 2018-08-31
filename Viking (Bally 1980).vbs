'Viking (Bally 1980) v1.0 by bord
'DOF by Arngrim

' Thalamus 2018-07-24
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Moved solenoids above table1_init
' Changed UseSolenoids=1 to 2
' No special SSF tweaks yet.
' Wob 2018-08-09
' Added vpmInit Me to table init and both cSingleLFlip and /cSingleRFlip
Option Explicit
Randomize

Const cGameName = "vikingb"

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "Can't open controller.vbs"
On Error Goto 0

LoadVPM "01130100", "Bally.VBS", 3.21  'Viking
Dim DesktopMode: DesktopMode = table1.ShowDT


'************************************************
'************************************************
'************************************************
'************************************************
'************************************************
Const UseSolenoids = 2
' Wob: Added for Fast Flips (No upper Flippers)
Const cSingleLFlip = 0
Const cSingleRFlip = 0
Const UseLamps = True
Const UseSync = False

' Standard Sounds
Const SSolenoidOn = "fx_solenoid"
Const SSolenoidOff = "fx_solenoidoff"
Const SCoin = "fx_coin"

Dim bsTrough, dtleft, dtright

Dim nfozzy: nfozzy=1            'set to 0 for standard vpx flipper physics or 1 for nfozzys adjustments to rate of return

'*************nFozzy flipper routine part 1
dim returnspeed, lfstep, rfstep
returnspeed = leftflipper.return
lfstep = 1
rfstep = 1

sub leftflipper_timer()
    select case lfstep
        Case 1: leftflipper.return = returnspeed * 0.6 :lfstep = lfstep + 1
        Case 2: leftflipper.return = returnspeed * 0.7 :lfstep = lfstep + 1
        Case 3: leftflipper.return = returnspeed * 0.8 :lfstep = lfstep + 1
        Case 4: leftflipper.return = returnspeed * 0.9 :lfstep = lfstep + 1
        Case 5: leftflipper.return = returnspeed * 1 :lfstep = lfstep + 1
        Case 6: leftflipper.timerenabled = 0 : lfstep = 1
    end select
end sub

sub rightflipper_timer()
    select case rfstep
        Case 1: rightflipper.return = returnspeed * 0.6 :rfstep = rfstep + 1
        Case 2: rightflipper.return = returnspeed * 0.7 :rfstep = rfstep + 1
        Case 3: rightflipper.return = returnspeed * 0.8 :rfstep = rfstep + 1
        Case 4: rightflipper.return = returnspeed * 0.9 :rfstep = rfstep + 1
        Case 5: rightflipper.return = returnspeed * 1 :rfstep = rfstep + 1
        Case 6: rightflipper.timerenabled = 0 : rfstep = 1
    end select
end sub

'*************nFozzy flipper routine part 1 end

Sub SolLFlipper(Enabled)
    If Enabled Then
        PlaySound SoundFX("fx_flipperup",DOFFlippers), 0, 1, -0.1, 0.25
        LeftFlipper.RotateToEnd
    Else
        PlaySound SoundFX("fx_flipperdown",DOFFlippers), 0, 1, -0.1, 0.25
        LeftFlipper.RotateToStart
'*************nFozzy flipper routine part 2
        if nfozzy=1 then
            LeftFlipper.TimerEnabled = 1
            LeftFlipper.TimerInterval = 16
            LeftFlipper.return = returnspeed * 0.5
        end if
'*************nFozzy flipper routine part 2 end
    End If
End Sub

Sub SolRFlipper(Enabled)
    If Enabled Then
        PlaySound SoundFX("fx_flipperup",DOFFlippers), 0, 1, 0.1, 0.25
        RightFlipper.RotateToEnd
   Else
        PlaySound SoundFX("fx_flipperdown",DOFFlippers), 0, 1, 0.1, 0.25
        RightFlipper.RotateToStart
'*************nFozzy flipper routine part 2b
        if nfozzy=1 then
            RightFlipper.TimerEnabled = 1
            RightFlipper.TimerInterval = 16
            RightFlipper.return = returnspeed * 0.5
        end if
'*************nFozzy flipper routine part 2b end
    End If
End Sub

Sub FlipperTimer_Timer
    'Add flipper, gate and spinner rotations here
    LFLogo.RotY = LeftFlipper.CurrentAngle
    RFLogo.RotY = RightFlipper.CurrentAngle
End Sub

'************************************************
' Solenoids
'************************************************
SolCallback(7) =    "bsTrough.SolOut"       'outhole kicker
'SolCallback(2) =   "Solknocker "               'knocker
SolCallback(8) =    "SolKickUpSaucer"       'kick up saucer
SolCallback(9) =    "SolKickDownSaucer"     'kick down saucer
'SolCallback(5) =   ""                  'left slingshot
'SolCallback(6) =   ""                  'right slingshot
SolCallback(1) =    "SolLeftTargetReset "           'in line drop target reset
SolCallback(2) =    "SolRightTargetReset "          '3 drop target reset
'SolCallback(9) =   "vpmSolSound "              'top left thumper bumper
'SolCallback(10) =  "vpmsolsound "              'top right thumper bumper
'SolCallback(11) =  "vpmSolSound "              'left side thumper bumper
'SolCallback(12) =  "vpmSolSound "              'right side thumper bumper
SolCallback(3) =    ""                  '3 drop target 1 (top)
SolCallback(4) =    ""                  '3 drop target 2
SolCallback(5) =    ""                  '3 drop target 3 (bottom)
'SolCallback(16) =  ""                  'coin lockout door
'SolCallback(17) =  ""                  'ki relay (flipper enable)
SolCallback(17) =   "SolTopSaucer"      'top saucer

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub Table1_Init
	 vpmInit Me
     With Controller
         .GameName = cGameName
         If Err Then MsgBox "Can't start Game " & cGameName & vbNewLine & Err.Description:Exit Sub
         .SplashInfoLine = ""
         .HandleKeyboard = 0
         .ShowTitle = 0
         .ShowDMDOnly = 1
         .ShowFrame = 0
         .HandleMechanics = False
		 .Hidden = 1
         On Error Resume Next
         .Run GetPlayerHWnd
         If Err Then MsgBox Err.Description
         On Error Goto 0
     End With

    PinMAMETimer.Interval = PinMAMEInterval
    PinMAMETimer.Enabled = 1


'Trough
    Set bsTrough=New cvpmBallStack
    with bsTrough
        .InitSw 0,8,0,0,0,0,0,0
        .InitKick BallRelease,90,7
        .InitExitSnd Soundfx("fx_ballrel",DOFContactors), Soundfx("fx_solenoid",DOFContactors)
        .Balls=1
    end with

     ' Nudging
     vpmNudge.TiltSwitch = 7
     vpmNudge.Sensitivity = 3
     vpmNudge.TiltObj = Array(leftslingshot, rightslingshot, sw37, sw38, sw39, sw40)

    Set dtLeft=New cvpmDropTarget
        dtLeft.InitDrop Array(sw1,sw2,sw3,sw4),Array(1,2,3,4)
        dtLeft.InitSnd SoundFX("fx2_droptarget",DOFDropTargets),SoundFX("fx2_DTReset",DOFDropTargets)

    Set dtRight=New cvpmDropTarget
        dtRight.InitDrop Array(sw17,sw18,sw19),Array(17,18,19)
        dtRight.InitSnd SoundFX("fx2_droptarget2",DOFDropTargets),SoundFX("fx2_DTReset",DOFDropTargets)

'*****Drop Lights Off
    dim xx

    For each xx in DTLeftLights: xx.state=0:Next
    For each xx in DTRightLights: xx.state=0:Next

    GILights 1

 End Sub

Sub GILights (enabled)
    Dim light
    For each light in GI:light.State = Enabled: Next
End Sub


Sub Table1_KeyDown(ByVal keycode)
    If keycode = LeftTiltKey Then Nudge 90, 2
    If keycode = RightTiltKey Then Nudge 270, 2
    If keycode = CenterTiltKey Then Nudge 0, 2

    If vpmKeyDown(keycode) Then Exit Sub
    If keycode = PlungerKey Then Plunger.PullBack: PlaySound "fx_plungerpull",0,1,0.25,0.25:    End If
End Sub

Sub Table1_KeyUp(ByVal keycode)
    If keycode = PlungerKey Then Plunger.Fire: PlaySound "fx_plunger",0,1,0.25,0.25
    If vpmKeyUp(keycode) Then Exit Sub

End Sub

Sub ShooterLane_Hit
End Sub

Sub BallRelease_UnHit
End Sub

Sub Drain_Hit()
    PlaySound "fx2_drain2",0,1,0,0.25 : bstrough.addball me
End Sub

Sub sw24_Hit   : Controller.Switch(24) = True : PlaySound "fx_hole-enter":End Sub
'
Sub sw32_Hit   : Controller.Switch(32) = True : PlaySound "fx_hole-enter":End Sub


Sub SolTopSaucer(enabled)
    if enabled then
        playsound Soundfx("fx_ballrel",DOFContactors)
        playsound Soundfx("fx_solenoid",DOFContactors)
        Controller.Switch(24) = false
        sw24.Kick  182, 15 + 5 * Rnd
    end if
End Sub

Sub SolKickUpSaucer(enabled)
    if enabled then
        playsound Soundfx("fx_ballrel",DOFContactors)
        playsound Soundfx("fx_solenoid",DOFContactors)
        Controller.Switch(32) = false
        sw32.Kick  10, 15 + 5 * Rnd
    end if
End Sub

Sub SolKickDownSaucer(enabled)
    if enabled then
        playsound Soundfx("fx_ballrel",DOFContactors)
        playsound Soundfx("fx_solenoid",DOFContactors)
        Controller.Switch(32) = false
        sw32.Kick  180, 15 + 5 * Rnd
    end if
End Sub

'Drop Targets
 Sub Sw1_Dropped:dtLeft.Hit 1 : GI_12b1.state=1 : GI_12a.state=1 : End Sub
 Sub Sw2_Dropped:dtLeft.Hit 2 : GI_12b.state=1 : End Sub
 Sub Sw3_Dropped:dtLeft.Hit 3 : GI_12c.state=1 : End Sub
 Sub Sw4_Dropped:dtLeft.Hit 4 : End Sub

 Sub Sw17_Dropped:dtRight.Hit 1 : End Sub
 Sub Sw18_Dropped:dtRight.Hit 2 : GI_14c.state=1 : End Sub
 Sub Sw19_Dropped:dtRight.Hit 3 : GI_13a.state=1 : End Sub

Sub SolRightTargetReset(enabled)
    dim xx
    if enabled then
        dtRight.SolDropUp enabled
        For each xx in DTRightLights: xx.state=0:Next
    end if
End Sub

Sub SolLeftTargetReset(enabled)
    dim xx
    if enabled then
        dtLeft.SolDropUp enabled
        For each xx in DTLeftLights: xx.state=0:Next
    end if
End Sub

'Bumpers

Sub sw37_Hit : vpmTimer.PulseSw 37 : playsound SoundFX("fx2_bumper_1",DOFContactors): End Sub
Sub sw38_Hit : vpmTimer.PulseSw 38 : playsound SoundFX("fx2_bumper_1",DOFContactors): End Sub
Sub sw39_Hit : vpmTimer.PulseSw 39 : playsound SoundFX("fx2_bumper_2",DOFContactors): End Sub
Sub sw40_Hit : vpmTimer.PulseSw 40 : playsound SoundFX("fx2_bumper_3",DOFContactors): End Sub

'Wire Triggers
Sub SW12_Hit:Controller.Switch(12)=1 : End Sub  'B
Sub SW12_unHit:Controller.Switch(12)=0:End Sub
Sub SW13_Hit:Controller.Switch(13)=1 : End Sub  'A
Sub SW13_unHit:Controller.Switch(13)=0:End Sub
Sub SW15_Hit:Controller.Switch(15)=1 : End Sub  'Side R.O. Button
Sub SW15_unHit:Controller.Switch(15)=0:End Sub
Sub SW26_Hit:Controller.Switch(26)=1 : End Sub  'Right out Rollover
Sub SW26_unHit:Controller.Switch(26)=0:End Sub
Sub SW27_Hit:Controller.Switch(27)=1 : End Sub  'Flip Feed Lane (Rt)
Sub SW27_unHit:Controller.Switch(27)=0:End Sub
Sub SW28_Hit:Controller.Switch(28)=1 : End Sub  'Flip Feed Lane (Lt)
Sub SW28_unHit:Controller.Switch(28)=0:End Sub
Sub SW29_Hit:Controller.Switch(29)=1 : End Sub  'Left outlane
Sub SW29_unHit:Controller.Switch(29)=0:End Sub
Sub SW30_Hit:Controller.Switch(30)=1 : End Sub  'Right Side Lane R.O.
Sub SW30_unHit:Controller.Switch(30)=0:End Sub

'Spinners
Sub sw14_Spin : vpmTimer.PulseSw (14) :PlaySound "fx_spinner": End Sub

'Targets
Sub sw5_Hit:vpmTimer.PulseSw (5):End Sub
Sub sw25_Hit:vpmTimer.PulseSw (25):End Sub

Sub SolKnocker(Enabled)
    If Enabled Then PlaySound SoundFX("Knocker",DOFKnocker)
End Sub

'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep

Sub RightSlingShot_Slingshot
    PlaySound SoundFX("fx_slingshot",DOFContactors), 0, 1, 0.05, 0.05
    vpmtimer.PulseSw(35)
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.TransZ = -20
    RStep = 0
    RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.TransZ = -10
        Case 4:RSLing2.Visible = 0:RSLing3.Visible = 1:sling1.TransZ = 0
        Case 5:RSLing3.Visible = 0:RSLing.Visible = 1:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
    PlaySound SoundFX("fx_slingshot",DOFContactors),0,1,-0.05,0.05
    vpmtimer.pulsesw(36)
    LSling.Visible = 0
    LSling1.Visible = 1
    sling2.TransZ = -20
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.TransZ = -10
        Case 4:LSLing2.Visible = 0:LSLing3.Visible = 1:sling2.TransZ = 0
        Case 5:LSLing3.Visible = 0:LSLing.Visible = 1:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub
'******************************
' Diverse Collection Hit Sounds
'******************************

'Sub aBumpers_Hit (idx): PlaySound SoundFX("fx_bumper", DOFContactors), 0, 1, pan(ActiveBall): End Sub
Sub aRollovers_Hit(idx):PlaySound "fx_sensor", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aGates_Hit(idx):PlaySound "fx_Gate", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aMetals_Hit(idx):PlaySound "fx_MetalHit", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aRubber_Bands_Hit(idx):PlaySound "fx_rubber", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aRubbers_Hit(idx):PlaySound "fx_postrubber", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aRubber_Posts_Hit(idx):PlaySound "fx_postrubber", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aRubber_Pins_Hit(idx):PlaySound "fx_postrubber", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aPlastics_Hit(idx):PlaySound "fx_PlasticHit", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aTargets_Hit(idx):PlaySound "fx2_target", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aWoods_Hit(idx):PlaySound "fx_Woodhit", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub

'-------------------------------------
' Map lights into array
' Set unmapped lamps to Nothing
'-------------------------------------
Set Lights(1)  = l1
Set Lights(2)  = l2
Set Lights(3)  = l3
Set Lights(4)  = l4
Set Lights(5)  = l5
Set Lights(6)  = l6
Set Lights(7)  = l7
Set Lights(8)  = l8
Set Lights(9)  = l9
Lights(10) = Array(l10,l10a)
'Set Lights(11) = l11 'Shoot Again
Set Lights(12) = l12
'Set Lights(13) = l13 'Ball In Play
Set Lights(14) = l14
Set Lights(15) = l15
'Set Lights(16) = unused
Set Lights(17) = l17
Set Lights(18) = l18
Set Lights(19) = l19
Set Lights(20) = l20
Set Lights(21) = l21
Set Lights(22) = l22
Set Lights(23) = l23
Set Lights(24) = l24
Set Lights(25) = l25
Set Lights(26) = l26
'Set Lights(27) = l27 'LightMatch
Set Lights(28) = l28
'Set Lights(29) = l29 'LightHighScore
Set Lights(30) = l30
Set Lights(31) = l31
'Set Lights(32) = unused
Set Lights(33) = l33
Set Lights(34) = l34
Set Lights(35) = l35
Set Lights(36) = l36
Set Lights(37) = l37
Set Lights(38) = l38
Set Lights(39) = l39
Set Lights(40) = l40
Set Lights(41) = l41
Set Lights(42) = l42
Set Lights(43) = l43
Set Lights(44) = l44
'Set Lights(45) = l45 'LightGameOver
Set Lights(46) = l46
Set Lights(47) = l47
'Set Lights(48) = unused
Set Lights(49) = l49
Set Lights(50) = l50
Set Lights(51) = l51
Set Lights(52) = l52
Set Lights(53) = l53
Set Lights(54) = l54
Set Lights(55) = l55
Set Lights(56) = l56
Set Lights(57) = l57
Set Lights(58) = l58
'Set Lights(59) = l59 'LightCredit
Set Lights(60) = l60
'Set Lights(61) = l61 'LightTilt
Set Lights(62) = l62
Set Lights(63) = l63

'Set Lights(10) = GILeft        'PF_Left
'Set Lights(42) = GIRight   'PF_Right
'Set Lights(26) = GICenter  'PF_Center

'---------------------------------------------------------------
' Edit the dip switches
'---------------------------------------------------------------
Sub editDips
    Dim vpmDips : Set vpmDips = New cvpmDips
    with vpmDips
        .AddForm 315, 370, "Viking DIP Switch Settings"
        .AddFrame 0,0,190,"Maximum credits",&H03000000,Array("10 credits",0,"20 credits",&H01000000,"30 credits",&H02000000,"40 credits",&H03000000)'dip 25&26
        .AddFrame 0,76,190,"High game to date",&H00300000,Array("no award",0,"1 credit",&H00100000,"2 credits",&H00200000,"3 credits",&H00300000)'dip 21&22
        .AddFrame 0,152,190,"Bumper points adjust",&H00800000,Array("100 points",0,"1.000 points",&H00800000)'dip 24
        .AddFrame 0,244,190,"Red target adjust",&H00000080,Array("does not add bonus",0,"adds 5 extra bonus",&H00000080)'dip 8
        .AddFrame 205,0,190,"Sound features",&H30000000,Array("chime effects",0,"noises and no background",&H10000000,"noise effects",&H20000000,"noises and background",&H30000000)'dip 29&30
        .AddFrame 205,76,190,"High score feature",&H00000060,Array("points",0,"extra ball",&H00000040,"replay",&H00000060)'dip 6&7
        .AddChk 205,155,190,Array("Match feature",&H08000000)'dip 28
'        .AddChk 205,175,190,Array("Credits displayed",0)'dip 27
		.AddChk 205,175,190,Array("Credits displayed",&H04000000)'dip 27
        .AddChk 205,195,190,Array("25K light in memory",&H00004000)'dip 15
        .AddChk 205,215,190,Array("3 bank drop target in memory",32768)'dip 16
        .AddChk 205,235,190,Array("Special and lock ball in memory",&H00002000)'dip 14
        .AddChk 205,255,190,Array("In-line extra ball && special in memory",&H80000000)'dip 32
        .AddChk 205,275,190,Array("In-line drop targets points in memory",&H00400000)'dip 23
        .AddLabel 50,300,300,20,"After hitting OK, press F3 to reset game with new settings."
        .ViewDips
    end with
End Sub
Set vpmShowDips = GetRef("editDips")

Sub table1_Paused:Controller.Pause = 1:End Sub
Sub table1_unPaused:Controller.Pause = 0:End Sub
Sub table1_Exit:Controller.Stop:End Sub

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
  Vol = Csng(BallVel(ball) ^2 / 2000)
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

Const tnob = 20 ' total number of balls
ReDim rolling(tnob)
InitRolling

Sub InitRolling
    Dim i
    For i = 0 to tnob
        rolling(i) = False
    Next
End Sub

Sub RollingUpdate()
    Dim BOT, b, ballpitch
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
  If Table1.VersionMinor > 3 OR Table1.VersionMajor > 10 Then
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 200, Pan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
  Else
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 200, Pan(ball1), 0, Pitch(ball1), 0, 0
  End if
End Sub


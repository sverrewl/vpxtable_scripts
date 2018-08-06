Option Explicit
Randomize

' Thalamus 2018-07-24
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' No special SSF tweaks yet.

Const cGameName = "pnkpnthr"

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "Can't open controller.vbs"
On Error Goto 0

LoadVPM "01210000", "sys80.VBS", 3.1
Set LampCallback = GetRef("UpdateMultipleLamps")

'**********************************************************
'********       OPTIONS     *******************************
'**********************************************************

Dim BallShadows: Ballshadows=1          '******************set to 1 to turn on Ball shadows
Dim FlipperShadows: FlipperShadows=1  '***********set to 1 to turn on Flipper shadows

'*********** Desktop/Cabinet settings ************************
dim HiddenValue
If Table1.ShowDT = true Then
    HiddenValue = 0
    SideRailLeft.visible=1
    SideRailRight.visible=1
Else
    HiddenValue = 1
    SideRailLeft.visible=1
    SideRailRight.visible=1
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

Dim bsTrough, bskicker, bslLock, bsrLock, dttop, dtleft, FastFlips

DisplayTimer.Enabled = true

Sub FlipperTimer_Timer
    'Add flipper, gate and spinner rotations here
   FlipperT2.RotY = LeftFlipper.CurrentAngle
    FlipperT5.RotY = RightFlipper.CurrentAngle
    FlipperT1.RotY = RightFlipper1.CurrentAngle
    FlipperLSh.rotz = LeftFlipper.currentangle '+ 45
   FlipperRSh.rotz = RightFlipper1.currentangle '+ 45
   FlipperRSh1.rotz = RightFlipper1.currentangle '+ 45
End Sub

Sub Table1_Init
     With Controller
         .GameName = cGameName
         If Err Then MsgBox "Can't start Game " & cGameName & vbNewLine & Err.Description:Exit Sub
         .SplashInfoLine = "Pink Panther (Gottlieb 1981)"&chr(13)&"1.0"
         .HandleKeyboard = 0
         .ShowTitle = 0
         .ShowDMDOnly = 1
         .ShowFrame = 0
         .HandleMechanics = False
     .Hidden = HiddenValue
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
        .InitSw 63,73,0,0,0,0,0,0
        .InitKick ballrelease, 110, 1
        .InitExitSnd Soundfx("fx_ballrel",DOFContactors), Soundfx("fx_solenoid",DOFContactors)
        .Balls=3
    end with

'Kickers
    Set bskicker=New cvpmBallStack
    with bskicker
        .InitSaucer sw13,13,255,12
        .InitExitSnd Soundfx("fx_ballrel",DOFContactors), Soundfx("fx_solenoid",DOFContactors)
        .kickanglevar=2
    end with

    Set bslLock=New cvpmBallStack
    with bslLock
        .InitSaucer sw43,43,-170,10
        .InitExitSnd Soundfx("fx_ballrel",DOFContactors), Soundfx("fx_solenoid",DOFContactors)
    end with

    Set bsrLock=New cvpmBallStack
    with bsrLock
        .InitSaucer sw53,53,320,11
        .InitExitSnd Soundfx("fx_ballrel",DOFContactors), Soundfx("fx_solenoid",DOFContactors)
    end with

' Nudging
    vpmNudge.TiltSwitch = 57
     vpmNudge.Sensitivity = 3
     vpmNudge.TiltObj = Array(leftslingshot, rightslingshot, bumper1, bumper2, bumper3, bumper4)

    Set dtTop=New cvpmDropTarget
        dtTop.InitDrop Array(sw51,sw61,sw71),Array(51,61,71)
        dtTop.InitSnd SoundFX("fx_droptarget",DOFDropTargets),SoundFX("fx_DTReset",DOFDropTargets)

    Set dtLeft=New cvpmDropTarget
        dtLeft.InitDrop Array(sw40,sw50,sw60,sw70),Array(40,50,60,70)
        dtLeft.InitSnd SoundFX("fx_droptarget",DOFDropTargets),SoundFX("fx_DTReset",DOFDropTargets)

if ballshadows=1 then
        BallShadowUpdate.enabled=1
      else
        BallShadowUpdate.enabled=0
    end if

    if flippershadows=1 then
        FlipperLSh.visible=1
        FlipperRSh.visible=1
        FlipperRSh1.visible=1
      else
        FlipperLSh.visible=0
        FlipperRSh.visible=0
        FlipperRSh1.visible=0
    end if

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
SolCallback(1) =    "SolKicker"
SolCallback(2) =    "SolTopTargetReset"
'SolCallback(3) =   coin1
'SolCallback(4) =   coin2
SolCallback(5) =    "SolLeftTargetReset"
SolCallback(6) =    "delayedSolOut"
'SolCallback(7) =   coin3
SolCallback(8) =    "solknocker"
SolCallback(9) =    "bstrough.SolIn"
SolCallback(10) = "FastFlips.TiltSol"

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
     If Enabled Then
         PlaySound SoundFX("fx_Flipperup",DOFContactors):LeftFlipper.RotateToEnd
     Else
         PlaySound SoundFX("fx_Flipperdown",DOFContactors):LeftFlipper.RotateToStart
     End If
  End Sub

Sub SolRFlipper(Enabled)
     If Enabled Then
         PlaySound SoundFX("fx_Flipperup",DOFContactors):RightFlipper.RotateToEnd:RightFlipper1.RotateToEnd
     Else
         PlaySound SoundFX("fx_Flipperdown",DOFContactors):RightFlipper.RotateToStart:RightFlipper1.RotateToStart
     End If
End Sub

'***** EM kicker animation

Dim RightEMPos

Sub SolKicker(enabled)
	If enabled Then
		bskicker.ExitSol_On
		RightEMpos = 0
		PkickarmR.RotZ = 2
		RightEMTimer.Enabled = 1
	End If
End Sub

Sub RightEMTimer_Timer
    Select Case RightEMpos
        Case 1:PkickarmR.Rotz = 15
        Case 2:PkickarmR.Rotz = 15
        Case 3:PkickarmR.Rotz = 15
        Case 4:PkickarmR.Rotz = 8
        Case 5:PkickarmR.Rotz = 4
        Case 6:PkickarmR.Rotz = 2
        Case 7:PkickarmR.Rotz = 0:RightEMTimer.Enabled = 0
    End Select
    RightEMpos = RightEMpos + 1
End Sub

Dim timer

Sub delayedSolOut(enabled)
    if not timer then bsTrough.SolOut enabled
    if not enabled then timer = True : ballrelease.TimerEnabled = True
End Sub

Sub ballrelease_Timer()
    ballrelease.TimerEnabled = False
    timer = False
End Sub

Sub UpdateMultipleLamps
    vpmNudge.solGameOn (Controller.Lamp(0) and not Controller.Lamp(1))
    if l8.state=1 then bsLLock.SolOut True
    if l9.state=1 then bsRLock.SolOut True
    L20a.State=L20.State
    L21a.State=L21.State
    L22a.State=L22.State
End Sub

'*****Drop Lights Off
   dim xx
    For each xx in dtLeftLights: xx.state=0:Next

'**********************************************************************************************************
'Plunger code
'**********************************************************************************************************

Sub Table1_KeyDown(ByVal KeyCode)
    If KeyCode = LeftFlipperKey then FastFlips.FlipL True :  FastFlips.FlipUL True
    If KeyCode = RightFlipperKey then FastFlips.FlipR True :  FastFlips.FlipUR True
    If KeyDownHandler(keycode) Then Exit Sub
    If keycode = PlungerKey Then Plunger.Pullback:playsound"plungerpull"
End Sub

Sub Table1_KeyUp(ByVal KeyCode)
    If KeyCode = LeftFlipperKey then FastFlips.FlipL False :  FastFlips.FlipUL False
    If KeyCode = RightFlipperKey then FastFlips.FlipR False :  FastFlips.FlipUR False
    If KeyUpHandler(keycode) Then Exit Sub
    If keycode = PlungerKey Then Plunger.Fire:PlaySound"plunger"
End Sub

Sub Drain_Hit():PlaySound "fx_drain":bstrough.addball me:End Sub
Sub sw13_Hit:PlaySound "kicker_enter_center":bskicker.AddBall 0:End Sub
Sub sw43_Hit:PlaySound "kicker_enter_center":bslLock.AddBall 0:End Sub
Sub sw53_Hit:PlaySound "kicker_enter_center":bsrLock.AddBall 0:End Sub

'Drop Targets
Sub Sw51_Dropped:dtTop.Hit 1 : End Sub
 Sub Sw61_Dropped:dtTop.Hit 2 : End Sub
 Sub Sw71_Dropped:dtTop.Hit 3 : End Sub

 Sub Sw40_Dropped:dtLeft.Hit 1 : D1L1.state=1 : End Sub
 Sub Sw50_Dropped:dtLeft.Hit 2 : D2L1.state=1 : D2L2.state=1 : End Sub
 Sub Sw60_Dropped:dtLeft.Hit 3 : D3L1.state=1 : D3L2.state=1 : End Sub
 Sub Sw70_Dropped:dtLeft.Hit 4 : D4L1.state=1 : D4L2.state=1 : End Sub

Sub SolTopTargetReset(enabled)
    dim xx
    if enabled then
        dtTop.SolDropUp enabled
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

Sub bumper1_Hit : vpmTimer.PulseSw 23 : playsound SoundFX("fx_bumper1",DOFContactors): DOF 206, DOFPulse:End Sub
Sub bumper2_Hit : vpmTimer.PulseSw 23 : playsound SoundFX("fx_bumper2",DOFContactors): DOF 209, DOFPulse:End Sub
Sub bumper3_Hit : vpmTimer.PulseSw 23 : playsound SoundFX("fx_bumper3",DOFContactors): DOF 208, DOFPulse:End Sub
Sub bumper4_Hit : vpmTimer.PulseSw 23 : playsound SoundFX("fx_bumper4",DOFContactors): DOF 207, DOFPulse:End Sub

'Wire Triggers
Sub SW02_Hit:Controller.Switch(02)=1:PlaySound"rollover":End Sub    'P
Sub SW02_unHit:Controller.Switch(02)=0:End Sub
Sub SW12_Hit:Controller.Switch(12)=1:PlaySound"rollover":End Sub    'I
Sub SW12_unHit:Controller.Switch(12)=0:End Sub
Sub SW22_Hit:Controller.Switch(22)=1:PlaySound"rollover":End Sub    'N
Sub SW22_unHit:Controller.Switch(22)=0:End Sub
Sub SW32_Hit:Controller.Switch(32)=1:PlaySound"rollover":End Sub    'K
Sub SW32_unHit:Controller.Switch(32)=0:End Sub
Sub SW03_Hit:Controller.Switch(03)=1:PlaySound"rollover":End Sub    'right inlane
Sub SW03_unHit:Controller.Switch(03)=0:End Sub
Sub SW42_Hit:Controller.Switch(42)=1:PlaySound"rollover":End Sub    'left inlane
Sub SW42_unHit:Controller.Switch(42)=0:End Sub
Sub SW62_Hit:Controller.Switch(62)=1:PlaySound"rollover":End Sub    'left outlane
Sub SW62_unHit:Controller.Switch(62)=0:End Sub
Sub SW72_Hit:Controller.Switch(72)=1:PlaySound"rollover":End Sub    'right outlane
Sub SW72_unHit:Controller.Switch(72)=0:End Sub

'Targets
Sub sw00_Hit:vpmTimer.PulseSw (00):playsound"target":End Sub
Sub sw10_Hit:vpmTimer.PulseSw (10):playsound"target":End Sub
Sub sw20_Hit:vpmTimer.PulseSw (20):playsound"target":End Sub
Sub sw30_Hit:vpmTimer.PulseSw (30):playsound"target":End Sub
Sub sw01_Hit:vpmTimer.PulseSw (01):playsound"target":End Sub
Sub sw11_Hit:vpmTimer.PulseSw (11):playsound"target":End Sub
Sub sw21_Hit:vpmTimer.PulseSw (21):playsound"target":End Sub
Sub sw31_Hit:vpmTimer.PulseSw (31):playsound"target":End Sub
Sub sw41_Hit:vpmTimer.PulseSw (41):playsound"target":End Sub

Sub SolKnocker(Enabled)
    If Enabled Then PlaySound SoundFX("Knocker",DOFKnocker)
End Sub

'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep, D1Step, D2Step, D3Step, D4Step, D5Step, D6Step

Sub RightSlingShot_Slingshot
    PlaySound SoundFX("fx_slingshot",DOFContactors), 0, 1, 0.05, 0.05
    DOF 202, DOFPulse
    vpmtimer.PulseSw(33)
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
    DOF 201, DOFPulse
    vpmtimer.pulsesw(33)
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

Sub Dingwall1_Slingshot
    vpmtimer.pulsesw(33)
    Rubber7a.Visible = 0
    Rubber7a1.Visible = 1
    D1Step = 0
    Dingwall1.TimerEnabled = 1
End Sub

Sub Dingwall1_Timer
    Select Case D1Step
        Case 3:Rubber7a1.Visible = 0:Rubber7a2.Visible = 1
        Case 4:Rubber7a2.Visible = 0:Rubber7a.Visible = 1:Dingwall1.TimerEnabled = 0
    End Select
    D1Step = D1Step + 1
End Sub

Sub Dingwall2_Slingshot
    Rubber7a.Visible = 0
    Rubber7a3.Visible = 1
    D2Step = 0
    Dingwall2.TimerEnabled = 1
End Sub

Sub Dingwall2_Timer
    Select Case D2Step
        Case 3:Rubber7a3.Visible = 0:Rubber7a4.Visible = 1
        Case 4:Rubber7a4.Visible = 0:Rubber7a.Visible = 1:Dingwall2.TimerEnabled = 0
    End Select
    D2Step = D2Step + 1
End Sub

Sub Dingwall3_Slingshot
    Rubberc.Visible = 0
    Rubberc1.Visible = 1
    D3Step = 0
    Dingwall3.TimerEnabled = 1
End Sub

Sub Dingwall3_Timer
    Select Case D3Step
        Case 3:Rubberc1.Visible = 0:Rubberc2.Visible = 1
        Case 4:Rubberc2.Visible = 0:Rubberc.Visible = 1:Dingwall3.TimerEnabled = 0
    End Select
    D3Step = D3Step + 1
End Sub

Sub Dingwall4_Slingshot
    Rubberd.Visible = 0
    Rubberd1.Visible = 1
    D4Step = 0
    Dingwall4.TimerEnabled = 1
End Sub

Sub Dingwall4_Timer
    Select Case D4Step
        Case 3:Rubberd1.Visible = 0:Rubberd2.Visible = 1
        Case 4:Rubberd2.Visible = 0:Rubberd.Visible = 1:Dingwall4.TimerEnabled = 0
    End Select
    D4Step = D4Step + 1
End Sub

Sub Dingwall5_Slingshot
    vpmtimer.pulsesw(33)
    Rubbere.Visible = 0
    Rubbere1.Visible = 1
    D5Step = 0
    Dingwall5.TimerEnabled = 1
End Sub

Sub Dingwall5_Timer
    Select Case D5Step
        Case 3:Rubbere1.Visible = 0:Rubbere2.Visible = 1
        Case 4:Rubbere2.Visible = 0:Rubbere.Visible = 1:Dingwall5.TimerEnabled = 0
    End Select
    D5Step = D5Step + 1
End Sub

Sub Dingwall6_Slingshot
    vpmtimer.pulsesw(33)
    Rubberf.Visible = 0
    Rubberf1.Visible = 1
    D6Step = 0
    Dingwall6.TimerEnabled = 1
End Sub

Sub Dingwall6_Timer
    Select Case D6Step
        Case 3:Rubberf1.Visible = 0:Rubberf2.Visible = 1
        Case 4:Rubberf2.Visible = 0:Rubberf.Visible = 1:Dingwall6.TimerEnabled = 0
    End Select
    D6Step = D6Step + 1
End Sub

'******************************
' Diverse Collection Hit Sounds
'******************************
'Sub aMetals_Hit(idx):PlaySound "fx_MetalHit", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
'Sub aRubber_Bands_Hit(idx):PlaySound "fx_rubber", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
'Sub aRubbers_Hit(idx):PlaySound "fx_postrubber", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
'Sub aRubber_Posts_Hit(idx):PlaySound "fx_postrubber", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
'Sub aRubber_Pins_Hit(idx):PlaySound "fx_postrubber", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
Sub aPlastics_Hit(idx):PlaySound "fx_PlasticHit", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
'Sub aTargets_Hit(idx):PlaySound "fx2_target", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
Sub aWoods_Hit(idx):PlaySound "fx_Woodhit", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub

Sub aRubber_Pins_Hit (idx)
    PlaySound "pinhit_low", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub aTargets_Hit (idx)
    PlaySound "target", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub aMetals_Hit (idx)
    PlaySound "metalhit_medium", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub aGates_Hit (idx)
    PlaySound "gate4", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub aRubber_Bands_Hit(idx)
    dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
    If finalspeed > 20 then
        PlaySound "fx_rubber2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    End if
    If finalspeed >= 6 AND finalspeed <= 20 then
        RandomSoundRubber()
    End If
End Sub

Sub aRubber_Posts_Hit(idx)
    dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
    If finalspeed > 16 then
        PlaySound "fx_postrubber", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    End if
    If finalspeed >= 6 AND finalspeed <= 16 then
        RandomSoundRubber()
    End If
End Sub

Sub RandomSoundRubber()
    Select Case Int(Rnd*3)+1
        Case 1 : PlaySound "rubber_hit_1", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
        Case 2 : PlaySound "rubber_hit_2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
        Case 3 : PlaySound "rubber_hit_3", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
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
        Case 1 : PlaySound "flip_hit_1", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
        Case 2 : PlaySound "flip_hit_2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
        Case 3 : PlaySound "flip_hit_3", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    End Select
End Sub

'GI uselamps workaround
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

'-------------------------------------
' Map lights into array
' Set unmapped lamps to Nothing
'-------------------------------------

Set Lights(0)  = l0 'ball in play
'Set Lights(1)  = l1 'tilt
set Lights(1) = GIlamps 'GI circuit
'Set Lights(2)  = l2    coin lockout coil
'Set Lights(3)  = l3    shoot again
Set Lights(8)  = l8 'left capture release
Set Lights(9)  = l9 'right capture release
'Set Lights(10) = l10   high score
'Set Lights(11) = l11   game over
Set Lights(12) = l12
Set Lights(13) = l13
Set Lights(14) = l14
Set Lights(15) = l15
Set Lights(16) = l16
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
Set Lights(27) = l27
Set Lights(28) = l28
Set Lights(29) = l29
Set Lights(30) = l30
Set Lights(31) = l31
Set Lights(32) = l32
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
Set Lights(45) = l45
Set Lights(46) = l46
Set Lights(47) = l47
Set Lights(48) = l48
Set Lights(49) = l49
Set Lights(50) = l50
Set Lights(51) = l51

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'           LL      EEEEEE  DDDD        ,,   SSSSS
'           LL      EE      DD  DD      ,,  SS
'           LL      EE      DD   DD      ,   SS
'           LL      EEEE    DD   DD            SS
'           LL      EE      DD  DD              SS
'           LLLLLL  EEEEEE  DDDD            SSSSS
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
'       7 Digit Array
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

Dim LED7(35)
LED7(0)=Array()
LED7(1)=Array()
LED7(2)=Array()
LED7(3)=Array()
LED7(4)=Array()
LED7(5)=Array()
LED7(6)=Array()
LED7(7)=Array()
LED7(8)=Array()
LED7(9)=Array()
LED7(10)=Array()
LED7(11)=Array()
LED7(12)=Array()
LED7(13)=Array()
LED7(14)=Array()
LED7(15)=Array()
LED7(16)=Array()
LED7(17)=Array()
LED7(18)=Array()
LED7(19)=Array()
LED7(20)=Array()
LED7(21)=Array()
LED7(22)=Array()
LED7(23)=Array()

LED7(24)=Array(d261,d262,d263,d264,d265,d266,d267,LXM,d268)
LED7(25)=Array(d271,d272,d273,d274,d275,d276,d277,LXM,d278)
LED7(26)=Array(d241,d242,d243,d244,d245,d246,d247,LXM,d248)
LED7(27)=Array(d251,d252,d253,d254,d255,d256,d257,LXM,d258)

LED7(28)=Array()
LED7(29)=Array()
LED7(30)=Array()
LED7(31)=Array()
LED7(32)=Array()
LED7(33)=Array()
LED7(34)=Array()
LED7(35)=Array()

Sub DisplayTimer_Timer
    Dim ChgLED, II, Num, Chg, Stat, Obj
    ChgLED = Controller.ChangedLEDs(0, &Hffffffff)
    If Not IsEmpty(ChgLED) Then
        For II = 0 To UBound(ChgLED)
            Num = ChgLED(II, 0):Chg = ChgLED(II, 1):Stat = ChgLED(II, 2)
            If Num > 23 Then
                For Each Obj In LED7(Num)
                    If Chg And 1 Then Obj.State = Stat And 1

                    Chg = Chg \ 2:Stat = Stat \ 2
                Next
            End If
        Next
    End If
End Sub


'*****************************************
'           BALL SHADOW
'*****************************************
Dim BallShadow
BallShadow = Array (BallShadow1,BallShadow2,BallShadow3)

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
            BallShadow(b).X = ((BOT(b).X) - (Ballsize/16) + ((BOT(b).X - (Table1.Width/2))/17))' + 13
       Else
            BallShadow(b).X = ((BOT(b).X) + (Ballsize/16) + ((BOT(b).X - (Table1.Width/2))/17))' - 13
       End If
        ballShadow(b).Y = BOT(b).Y + 10
        If BOT(b).Z > 20 Then
            BallShadow(b).visible = 1
        Else
            BallShadow(b).visible = 0
        End If
    Next
End Sub

'Gottlieb Pink Panther
'added by Inkochnito
Sub editDips
    Dim vpmDips : Set vpmDips = New cvpmDips
    With vpmDips
        .AddForm 700,400,"Pink Panther - DIP switches"
        .AddFrame 2,10,190,"Maximum credits",49152,Array("8 credits",0,"10 credits",32768,"15 credits",&H00004000,"25 credits",49152)'dip 15&16
       .AddFrame 2,86,190,"Coin chute 1 and 2 control",&H00002000,Array("Seperate",0,"Same",&H00002000)'dip 14
       .AddFrame 2,132,190,"Playfield special",&H00200000,Array("Replay",0,"Extra Ball",&H00200000)'dip 22
       .AddFrame 2,178,190,"Maximum blue diamond total",&H80000000,Array("maximum 40",0,"maximum 50",&H80000000)'dip32
       .AddFrame 2,224,190,"High score to date awards",&H00C00000,Array("Not displayed and no award",0,"Displayed and no award",&H00800000,"Displayed and 2 replays",&H00400000,"Displayed and 3 replays",&H00C00000)'dip 23&24
       .AddChk 2,300,190,Array("Sound when scoring?",&H01000000)'dip 25
       .AddChk 2,315,190,Array("Replay button tune?",&H02000000)'dip 26
       .AddChk 2,330,190,Array("Coin switch tune?",&H04000000)'dip 27
       .AddChk 2,345,190,Array("Credits displayed?",&H08000000)'dip 28
       .AddChk 2,360,190,Array("Match feature",&H00020000)'dip 18
       .AddChk 2,375,190,Array("Attract features",&H20000000)'dip 30
       .AddFrameExtra 205,10,190,"Attract tune",&H0200,Array("No attract tune",0,"attract tune played every 5 minutes",&H0200)'S-board dip 2
       .AddFrame 205,56,190,"Balls per game",&H00010000,Array("5 balls",0,"3 balls",&H00010000)'dip 17
       .AddFrame 205,102,190,"Replay limit",&H00040000,Array("No limit",0,"One per ball",&H00040000)'dip 19
       .AddFrame 205,148,190,"Novelty",&H00080000,Array("Normal",0,"Points",&H0080000)'dip 20
       .AddFrame 205,194,190,"Game mode",&H00100000,Array("Replay",0,"Extra ball",&H00100000)'dip 21
       .AddFrame 205,240,190,"3rd coin chute credits control",&H00001000,Array("No effect",0,"Add 9",&H00001000)'dip 13
       .AddFrame 205,286,190,"Tilt penalty",&H10000000,Array("Game over",0,"Ball in play",&H10000000)'dip 29
       .AddFrame 205,332,190,"Playfield special adjust",&H40000000,Array("On 20% longer than conservative",0,"Conservative",&H40000000)'dip 31
       .AddLabel 50,400,300,20,"After hitting OK, press F3 to reset game with new settings."
    End With
    Dim extra
    extra = Controller.Dip(4) + Controller.Dip(5)*256
    extra = vpmDips.ViewDipsExtra(extra)
    Controller.Dip(4) = extra And 255
    Controller.Dip(5) = (extra And 65280)\256 And 255
End Sub
Set vpmShowDips = GetRef("editDips")

Sub table1_Paused:Controller.Pause = 1:End Sub
Sub table1_unPaused:Controller.Pause = 0:End Sub

    Sub table1_Exit:Controller.Stop:End Sub



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
'    JP's VP10 Rolling Sounds
'*****************************************

Const tnob = 3 ' total number of balls
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
  If Table1.VersionMinor > 3 OR Table1.VersionMajor > 10 Then
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 200, Pan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
  Else
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 200, Pan(ball1), 0, Pitch(ball1), 0, 0
  End if
End Sub


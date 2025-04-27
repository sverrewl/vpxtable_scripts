' Thalamus - seems ok
Option Explicit
Randomize

Const BallSize = 50
Const BallMass = 1.1

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Dim x, bsTrough, dtL, dtM, bsLE, bsRI, bsBO

Const cGameName = "che_cho"

Const UseSolenoids = 2 'Fastflips enabled
Const UseLamps = 0
Const UseGI = 0
Const UseSync = 0 'set it to 1 if the table runs too fast
Const HandleMech = 0

Dim VarHidden, UseVPMDMD
If Table1.ShowDT = true then
    UseVPMDMD = True
    VarHidden = 1
    For each x in aReels
        x.Visible = 1
    Next
else
    UseVPMDMD = False
    VarHidden = 0
    For each x in aReels
        x.Visible = 0
    Next
    lrail.Visible = 0
    rrail.Visible = 0
end if

if B2SOn = true then VarHidden = 1

LoadVPM "01550000", "wpc.vbs", 3.26

' Standard Sounds
Const SSolenoidOn = "fx_solenoidon"
Const SSolenoidOff = "fx_solenoidoff"
Const SCoin = "fx_Coin"

'************
' Table init.
'************

Dim startgamesound

Sub table1_Init
    vpmInit me
    With Controller
        .GameName = cGameName
        If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description:Exit Sub
        .SplashInfoLine = "Cheech & Chong: Road-Trip'pin (Bally 2021)" & vbNewLine & ""
        .HandleKeyboard = 0
        .ShowTitle = 0
        .ShowDMDOnly = 1
        .ShowFrame = 0
        .HandleMechanics = 0
        .Hidden = VarHidden
        .Games(cGameName).Settings.Value("sound") = 1
        .Games(cGameName).Settings.Value("rol") = 0 '1= rotated display, 0= normal
        '.SetDisplayPosition 0,0, GetPlayerHWnd 'restore dmd window position
        On Error Resume Next
        Controller.SolMask(0) = 0
        vpmTimer.AddTimer 1000, "Controller.SolMask(0)=&Hffffffff'" 'ignore all solenoids - then add the Timer to renable all the solenoids after 2 seconds
        Controller.Run GetPlayerHWnd
        On Error Goto 0
    End With

    ' Nudging
    vpmNudge.TiltSwitch = 14
    vpmNudge.Sensitivity = 3
    vpmNudge.TiltObj = Array(Bumper1, Bumper2, Bumper3, LeftSlingshot, RightSlingshot, RightSlingShotTop)

    ' Trough
    set bsTrough = new cvpmBallStack
    bsTrough.InitSw 15, 16, 17, 18, 0, 0, 0, 0
    bsTrough.InitKick BallRelease, 45, 7
    bsTrough.InitExitSnd SoundFX("fx_ballrel", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
    bsTrough.Balls = 3

    ' Top droptargets
    Set dtL = New cvpmDropTarget
    dtL.InitDrop Array(sw51, sw52, sw53), Array(51, 52, 53)
    dtL.InitSnd SoundFX("", DOFDropTargets), SoundFX("fx_resetdrop", DOFContactors)
    dtL.CreateEvents "dtL"

    ' Mid droptargets
    Set dtM = New cvpmDropTarget
    dtM.InitDrop Array(sw54, sw55, sw56), Array(54, 55, 56)
    dtM.InitSnd SoundFX("", DOFDropTargets), SoundFX("fx_resetdrop", DOFContactors)
    dtM.CreateEvents "dtM"

    ' Left Eject
    Set bsLE = New cvpmBallStack
    bsLE.InitSaucer sw36, 36, 90, 20
    bsLE.InitExitSnd SoundFX("fx_kicker", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
    bsLE.KickAngleVar = 3
    bsLE.KickForceVar = 3

    ' Right Eject
    Set bsRI = New cvpmBallStack
    bsRI.InitSaucer sw35, 35, 0, 27
    bsRI.InitExitSnd SoundFX("fx_kicker", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
    bsRI.KickAngleVar = 3
    bsRI.KickForceVar = 3

    ' Bottom Eject
    Set bsBO = New cvpmBallStack
    bsBO.InitSaucer sw28, 28, 243, 24
    bsBO.InitExitSnd SoundFX("fx_kicker", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
    bsBO.KickZ = 1.5
    bsBO.KickAngleVar = 3
    bsBO.KickForceVar = 3

    ' Main Timer init
    PinMAMETimer.Interval = PinMAMEInterval
    PinMAMETimer.Enabled = 1

    startgamesound = 1

    PlaySound "", -1, .05 'adjust the volume using the second parameter.
End Sub

'Road Closed Drop Target

Sub RoadClosed_Hit
    RoadClosed.TimerEnabled = 1
    vpmTimer.PulseSw 37
    vpmTimer.PulseSw 38
    PlaySound SoundFX("fx_droptarget", DOFDropTargets), 0, 1, pan(ActiveBall)
    PlaySound "fx_horn"
End Sub

Sub RoadClosed_Timer
    RoadClosed.IsDropped = 0
    RoadClosed.TimerEnabled = 0
    PlaySound "fx_resetdrop"
End Sub

Sub table1_Paused:Controller.Pause = 1:End Sub
Sub table1_unPaused:Controller.Pause = 0:End Sub
Sub table1_exit:Controller.stop:End Sub

Sub table1_KeyDown(ByVal Keycode)
    If keycode = LeftTiltKey Then Nudge 90, 5:PlaySound SoundFX("fx_nudge", 0), 0, 1, -0.1, 0.25
    If keycode = RightTiltKey Then Nudge 270, 5:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0.1, 0.25
    If keycode = CenterTiltKey Then Nudge 0, 6:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0, 0.25
    If keycode = PlungerKey Then PlaySoundAtVol "fx_PlungerPull", Plunger, 1:Plunger.Pullback:PlaySound "fx_idlerev"

    If KeyCode = startgamekey and startgametimer.enabled = 0 and startgamesound = 1 then
        startgamesound = 0
        startgametimer.enabled = 1
    end if

    If vpmKeyDown(keycode)Then Exit Sub
End Sub

Sub table1_KeyUp(ByVal Keycode)
    If vpmKeyUp(keycode)Then Exit Sub
    If keycode = PlungerKey Then PlaySoundAtVol "fx_plunger", Plunger, 1:Plunger.Fire
End Sub

'***************
'  Clown Shake
'***************
Dim cBall
Set cBall = ckicker.createball
ckicker.Kick 0, 0

Sub ClownShake
    cball.velx = - 0.65 + RND(1)
    cball.vely = - 0.65 + RND(1)
End Sub

Sub ClownUpdate
    Dim a, b
    a = (ckicker.y - cball.y) / 2
    b = (cball.x - ckicker.x) / 2
    clown1.rotx = a
    clown1.rotz = b
    clownspring.rotx = a
    clownspring.rotz = b
    clownshadow.rotx = a
    clownshadow.rotz = b
    clownwires.rotx = a
    clownwires.rotz = b
    clownjointlit1.rotx = a
    clownjointlit1.rotz = b
End Sub

'*********
' Switches
'*********

' Slings
Dim LStep, RStep

Sub LeftSlingShot_Slingshot
    PFlmheadlightLeft.duration 0, 200, 1
    PFlmheadlightRight.duration 0, 200, 1
    lightf25.duration 1, 200, 0
    PlaySoundAtVol SoundFX("fx_slingshot", DOFContactors), ActiveBall, 1
    LeftSling4.Visible = 1
    Lemk.RotX = 26
    LStep = 0
    vpmTimer.PulseSw 37
    LeftSlingShot.TimerEnabled = 1
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
    PFlmheadlightLeft.duration 0, 200, 1
    PFlmheadlightRight.duration 0, 200, 1
    lightf26.duration 1, 200, 0
    PlaySoundAtVol SoundFX("fx_slingshot", DOFContactors), ActiveBall, 1
    RightSling4.Visible = 1
    Remk.RotX = 26
    RStep = 0
    vpmTimer.PulseSw 38
    RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 1:RightSLing4.Visible = 0:RightSLing3.Visible = 1:Remk.RotX = 14
        Case 2:RightSLing3.Visible = 0:RightSLing2.Visible = 1:Remk.RotX = 2
        Case 3:RightSLing2.Visible = 0:Remk.RotX = -20:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

Sub RightSlingShotTop_Slingshot
    if LoveMachine.rotz = 0 then
        RightSlingShotTop.uservalue = -1
        LoveMachineTimer.Interval = 10
        LoveMachineTimer.Enabled = 1
    end if
    lightf25.duration 1, 200, 0
    lightf26.duration 1, 200, 0
    PlaySoundAtVol SoundFX("fx_lowrider", DOFContactors), ActiveBall, 1
    RightSling8.Visible = 1
    RemkTop.RotX = 26
    RStep = 0
    vpmTimer.PulseSw 37
    vpmTimer.PulseSw 38
    RightSlingShotTop.TimerEnabled = 1
    ClownShake
End Sub

Sub RightSlingShotTop_Timer
    Select Case RStep
        Case 1:RightSLing8.Visible = 0:RightSLing7.Visible = 1:RemkTop.RotX = 14
        Case 2:RightSLing7.Visible = 0:RightSLing6.Visible = 1:RemkTop.RotX = 2
        Case 3:RightSLing6.Visible = 0:RemkTop.RotX = -20:RightSlingShotTop.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

Sub swtroughlighton_hit
    troughlight.state = 1
End Sub

Sub swtroughlightoff_hit
    troughlight.state = 0
Dim bulb
    For each bulb in aGIlights
       bulb.state = 1
   Next
End Sub

Sub factorylightsw_hit
    factorylight.duration 1, 5500, 0
End Sub

Sub swstrawhouselightsoff_hit
    strawhouselightA.duration 0, 100, 0
    strawhouselightB.duration 0, 100, 0
    PlaySoundAtVol "fx_balldrophard", ActiveBall, 1
End Sub

Sub swstrawhouselightsoff2_hit
    strawhouselightA.duration 0, 100, 0
    strawhouselightB.duration 0, 100, 0
End Sub

Sub billboardlightson_hit
    billboardlight1.state = 1
    billboardlight2.state = 1
    billboardlight3.state = 1
End Sub

Sub swpeterrooter_hit
me.uservalue=1
me.timerenabled= 1
Dim w
  w = INT(3 * RND(1) )
  Select Case w
  Case 0:PlaySound"fx_peterrooter"
  Case 1:PlaySound"fx_peterrooter"
    Case 2:PlaySound""
End Select
End Sub

Sub swvanwindowson_hit
    vanwindowslight.duration 1, 5200, 0
End Sub

Sub culdesacexit_hit
    PlaySound "fx_culdesacspinshort"
    StopSound "fx_culdesacspin":StopSound "fx_cheechdonuts":StopSound "fx_chongdonuts"
    PlaySound "fx_culdesacballdrop"
    If Table1.ShowDT then
    loopleftlightDT.duration 2, 950, 0
    else
    loopleftlightCAB.duration 2, 950, 0
    end If
end Sub

Sub swscooplightRIGHTon_hit
    scooplightright.state = 1
    boxlight.duration 1, 800, 0
    ClownShake
End Sub

Sub slyman_hit
me.uservalue=1
me.timerenabled= 1
Dim w
  w = INT(2 * RND(1) )
  Select Case w
  Case 0:PlaySound""
  Case 1:PlaySound"fx_ihavehairpiece":gi6.duration 2, 1000, 1:gi42.duration 2, 1000, 1:gi001.duration 2, 1000, 1
End Select
End Sub

Sub habib_hit
me.uservalue=1
me.timerenabled= 1
Dim w
  w = INT(2 * RND(1) )
  Select Case w
  Case 0:PlaySound""
  Case 1:PlaySound"fx_keeltheball":gi7.duration 2, 1000, 1:gi61.duration 2, 1000, 1:gi002.duration 2, 1000, 1
End Select
End Sub

Sub swSisterMaryElephant_hit
if Activeball.vely > 0 then
 Dim x
  x = INT(8 * RND(1) )
  Select Case x
  Case 0:PlaySound"fx_Class":SisterMaryLight.duration 2, 1300, 0
  Case 1:PlaySound"fx_Shutup":SisterMaryLight.duration 2, 1800, 0
    Case 2:PlaySound""
    Case 3:PlaySound""
    Case 4:PlaySound""
    Case 5:PlaySound""
    Case 6:PlaySound""
    Case 7:PlaySound"fx_Class":SisterMaryLight.duration 2, 1300, 0
End Select
End If
End Sub

Sub swBlindMelon_hit
if Activeball.vely > 0 then
 Dim y
  y = INT(8 * RND(1) )
  Select Case y
  Case 0:PlaySound"fx_ahyeaow":BlindMelonLight.duration 2, 1150, 0
  Case 1:PlaySound"fx_goindowntown":BlindMelonLight.duration 2, 2850, 0
    Case 2:PlaySound""
    Case 3:PlaySound""
    Case 4:PlaySound""
    Case 5:PlaySound""
    Case 6:PlaySound""
    Case 7:PlaySound"fx_ahyeaow":BlindMelonLight.duration 2, 1150, 0
End Select
End If
End Sub

Sub swstrawberryshouselit_Hit
        strawhouselightA.duration 1, 600, 0
        strawhouselightB.duration 1, 600, 0
end sub

Sub swfactorylightlit_Hit
        factorylight.duration 1, 600, 0
end sub

Sub swCuldesacSpinStart_Hit
        Playsound "fx_culdesacspin"
end Sub

Sub swCuldesacSpinStop_Hit
if Activeball.vely > 0 then
        Stopsound "fx_culdesacspin":StopSound "fx_cheechdonuts":StopSound "fx_chongdonuts"
End If
end Sub

Sub swCheckpointExitBallDrop_Hit
        Playsound "fx_CheckpointExitBallDrop"
end Sub

Sub swburnout_hit
if Activeball.vely < 0 then
 Dim x
  x = INT(6 * RND(1) )
  Select Case x
    Case 0:PlaySound"fx_burnout"
  Case 1:PlaySound"fx_burnout"
    Case 2:PlaySound"fx_burnout"
    Case 3:PlaySound"fx_burnout"
  Case 4:PlaySound"fx_burnout2"
    Case 5:PlaySound"fx_burnoutC&C"
End Select
End If
End Sub

Sub swCuldesacSpinCorC_hit
if Activeball.vely < 0 then
 Dim z
  z = INT(4 * RND(1) )
  Select Case z
  Case 0:PlaySound"fx_cheechdonuts"
  Case 1:PlaySound""
    Case 2:PlaySound"fx_chongdonuts"
    Case 3:PlaySound""
End Select
End If
End Sub

Sub drainpostrubber_Hit
        DrainpostLEDglow.duration 1, 200, 0
        drainpostapronreflection.duration 1, 200, 0
end sub

Sub swimpalaleft_Hit
if Activeball.vely < 0 then
        Playsound "fx_impalaleft"
End If
end Sub

Sub swimpalaright_Hit
if Activeball.vely < 0 then
        Playsound "fx_impalaright"
End If
end Sub

Sub swbottomleftrubber_Hit
        Playsound "fx_ricochet"
end Sub

Sub swbottomrightrubber_Hit
        Playsound "fx_ricochet"
end Sub

'**********************
'     Animations
'**********************

' Love Machine
Sub LoveMachineTimer_timer
    LoveMachine.rotz = LoveMachine.rotz + (RightSlingShotTop.uservalue * 0.9)
    if LoveMachine.rotz <- 15 Then
        RightSlingShotTop.uservalue = 1
    end If
    if LoveMachine.rotz> 0 Then
        LoveMachine.rotz = 0
        LoveMachineTimer.Enabled = 0
    end if
    LoveMachineHeadlights.rotz = LoveMachine.rotz
    LoveMachinePlate.rotz = LoveMachine.rotz
    LoveMachineShadow.rotx = LoveMachine.rotz
end sub

'Van
Sub swVan_hit
    PlaySound "fx_vanservo"
    If Van.transx = 0 then
        me.uservalue = 1
        me.timerenabled = 1
        me.timerinterval = 12
    End If
end sub

Sub swVan_timer
    if me.uservalue <12 then
        Van.transx = Van.transx - 5.94
    elseif me.uservalue <40 then
        Van.transx = Van.transx - 5.94
    elseif me.uservalue <50 then
        Van.transx = Van.transx - 5.94
    elseif me.uservalue <150 then
        if me.uservalue = 160 then PlaySound ""
        if me.uservalue = 170 then gi53.Duration 1, 3300, 0
        if me.uservalue = 180 then gi53.Duration 1, 2000, 0
        elseif me.uservalue <190 then
            Van.transx = Van.transx + 1.3
        elseif me.uservalue <200 then
            Van.transx = Van.transx + 1.5
        else
            Van.transx = Van.transx + 1
            if Van.transx> 0 Then
                Van.transx = 0
                me.timerenabled = 0
            End If
    End If
    vanbackshadow.transx = Van.transx
    vanbottomshadow.transx = Van.transx
    vanwindows.transx = Van.transx
    vanventlight.transx = Van.transx
    vanheadandtaillights.transx = Van.transx
    me.uservalue = me.uservalue + 1
end sub

'Van Kickers
Sub VanKickerEnter_Hit
    VanKickerEnter.destroyball
    VanKickerEnter.TimerEnabled = 1
End Sub

Sub VanKickerEnter_Timer
    VanKickerExit.createball
    VanKickerEnter.TimerEnabled = 0
    PlaySoundAtVol "fx_balldrop", VanKickerEnter, 1
End Sub

'Motorcycle Cop
Dim Mpos, MDir
Mpos = 0

Sub swMotorcycleCop_hit
    If Mpos = 0 then
        MDir = 5
        Me.timerinterval = 10
        vpmtimer.addtimer 2000, "swMotorcycleCop.timerenabled= 1 '"
    End If
end sub

Sub StartRightAnimations
    SaySomething = False
    StartHeadAnimations2
End Sub

Sub swMotorcycleCop_timer
    if Mpos = 600 then gi53.Duration 1, 3300, 0:PlaySound ""
    If Mpos = 10 AND Mdir = 5 Then StartRightAnimations:PlaySound "fx_itsacop"
    Mpos = Mpos + MDir
    If MPos> 725 Then MPos = 725:Me.TimerEnabled = 0
    RunnerRedLED.State = 2:RunnerRedLEDa.State = 2:RunnerRedLEDb.State = 2:RunnerRedLEDc.State = 2:RunnerRedLEDd.State = 2:RunnerRedLEDe.State = 2:RunnerBlueLED.State = 2:RunnerBlueLEDa.State = 2:RunnerBlueLEDb.State = 2:RunnerBlueLEDc.State = 2:RunnerBlueLEDd.State = 2:RunnerBlueLEDe.State = 2
    If Mpos <= 0 Then Mpos = 0:RunnerRedLED.State = 0:RunnerRedLEDa.State = 0:RunnerRedLEDb.State = 0:RunnerRedLEDc.State = 0:RunnerRedLEDd.State = 0:RunnerRedLEDe.State = 0:RunnerBlueLED.State = 0:RunnerBlueLEDa.State = 0:RunnerBlueLEDb.State = 0:RunnerBlueLEDc.State = 0:RunnerBlueLEDd.State = 0:RunnerBlueLEDe.State = 0:Me.TimerEnabled = 0
    PlaySound "fx_motorcycleservo"
    MotorcycleCop.transx = Mpos
    MotorcycleCopScrewA.transx = Mpos
    MotorcycleCopScrewB.transx = Mpos
    MotorcycleCopBracket.transx = Mpos
end sub

Sub sw28_Unhit:MDir = -2.5:RunnerRedLED.State = 0:RunnerRedLEDa.State = 0:RunnerRedLEDb.State = 0:RunnerRedLEDc.State = 0:RunnerRedLEDd.State = 0:RunnerRedLEDe.State = 0:RunnerBlueLED.State = 0:RunnerBlueLEDa.State = 0:RunnerBlueLEDb.State = 0:RunnerBlueLEDc.State = 0:RunnerBlueLEDd.State = 0:RunnerBlueLEDe.State = 0:swMotorcycleCop.timerenabled = 1:End Sub 'restart the animation after the ball is kicked from sw28

' Bumpers
Sub Bumper1_Hit:Playsound "fx_disc1spin":vpmTimer.PulseSw 26:PlaySoundAtVol SoundFX("fx_bumper", DOFContactors), ActiveBall, 1:gi1.duration 0, 150, 1:RotDisc1Step = 7:Disc1Timer.Enabled = 1:ClownShake:End Sub
Sub Bumper2_Hit:Playsound "fx_disc2spin":vpmTimer.PulseSw 25:PlaySoundAtVol SoundFX("fx_bumper", DOFContactors), ActiveBall, 1:gi17.duration 0, 150, 1:RotDisc2Step = 7:Disc2Timer.Enabled = 1:ClownShake:End Sub
Sub Bumper3_Hit:Playsound "fx_disc3spin":vpmTimer.PulseSw 27:PlaySoundAtVol SoundFX("fx_bumper", DOFContactors), ActiveBall, 1:gi3.duration 0, 150, 1:RotDisc3Step = 7:Disc3Timer.Enabled = 1:ClownShake:End Sub

' Bumper Discs
Dim RotDisc1, RotDisc1Step
RotDisc1 = 0
Sub Disc1Timer_Timer
    RotDisc1 = (RotDisc1 + RotDisc1Step)MOD 360
    disc1.RotY = RotDisc1
    disc1screw.RotZ = RotDisc1
    RotDisc1Step = RotDisc1Step - 0.05
    If RotDisc1Step <1 Then Disc1Timer.Enabled = 0
End Sub

Dim RotDisc2, RotDisc2Step
RotDisc2 = 0
Sub Disc2Timer_Timer
    RotDisc2 = (RotDisc2 - RotDisc2Step)MOD 360
    disc2.RotY = RotDisc2
    disc2screw.RotZ = RotDisc2
    RotDisc2Step = RotDisc2Step - 0.05
    If RotDisc2Step <1 Then Disc2Timer.Enabled = 0
End Sub

Dim RotDisc3, RotDisc3Step
RotDisc3 = 0
Sub Disc3Timer_Timer
    RotDisc3 = (RotDisc3 - RotDisc3Step)MOD 360
    disc3.RotY = RotDisc3
    disc3screw.RotZ = RotDisc3
    RotDisc3Step = RotDisc3Step - 0.05
    If RotDisc3Step <1 Then Disc3Timer.Enabled = 0
End Sub

' CulDeSac
Sub CulDeSac1_Hit:vpmTimer.PulseSw 27:culdesac1light.duration 1, 100, 0:End Sub
Sub CulDeSac2_Hit:vpmTimer.PulseSw 27:culdesac2light.duration 1, 100, 0:End Sub
Sub CulDeSac3_Hit:vpmTimer.PulseSw 27:culdesac3light.duration 1, 100, 0:End Sub
Sub CulDeSac4_Hit:vpmTimer.PulseSw 27:culdesac4light.duration 1, 100, 0:End Sub

' Drain & Saucers
Sub Drain_Hit:PlaySoundAtVol "fx_drain", Drain, 1:bsTrough.AddBall Me:
    gi66.duration 0, 3675, 1
    startgamesound = 1
    ReturnHeads
End Sub

Sub sw28_Hit:PlaySoundAtVol "fx_kicker_enter", ActiveBall, 1:bsBO.AddBall 0:End Sub
Sub sw36_Hit:PlaySoundAtVol "fx_kicker_enter", ActiveBall, 1:bsLE.AddBall 0:End Sub
Sub sw35_Hit:PlaySoundAtVol "fx_kicker_enter", ActiveBall, 1:bsRI.AddBall 0:billboardlight1.state = 0:billboardlight2.state = 0:billboardlight3.state = 0:End Sub

Sub CandCprimsReset_Hit
    ReturnHeads
End Sub

' Added Kickers
Sub kicker1_Hit
    PlaySoundAtVol "fx_hole_enter", ActiveBall, 1
    If bsLE.Balls Then
        vpmtimer.AddTimer 1000, "PlaySound""fx_popper"": kicker1.kick 165, 35, 1.5 '"
    Else
        kicker1.Destroyball
        sw36.CreateBall
        strawhouselightA.state = 1
        strawhouselightB.state = 1
        scooplightleft.duration 0, 2300, 1
        bsLE.AddBall 0
        SaySomething = False
        StartHeadAnimations
    End If
End Sub

Sub sw36_UnHit:RotateLeftHead.Enabled = 0:RotateRightHead.Enabled = 0:End Sub

' rotations of heads at sw36 and sw28
Dim LeftDir, LeftStep
Dim RightDir, RightStep
Dim SaySomething
Dim WhatToSay

Sub StartHeadAnimations
    LeftDir = -1
    LeftStep = 0
    RightDir = 1.5
    RightStep = 0
    RotateLeftHead.Interval = 2500
    RotateLeftHead.Enabled = 1
End Sub

Sub StartHeadAnimations2
    LeftDir = -1
    LeftStep = 0
    RightDir = 1.5
    RightStep = 0
    RotateLeftHead.Interval = 10
    RotateLeftHead.Enabled = 1
End Sub

Sub RotateLeftHead_Timer
    cheechspotlight.state = 1
    cheechspotlightglow.state = 1
    LeftStep = LeftStep + LeftDir
    RotateLeftHead.Interval = 10
    If LeftStep <-40 Then
        LeftStep = -40
        LeftDir = 1
        If SaySomething Then
            WhatToSay = INT(RND(1) * 6)
            Select Case WhatToSay
                Case 0:PlaySound "":RotateLeftHead.Interval = 4000:RotateRightHead.Interval = 2000
                Case 1:PlaySound "":RotateLeftHead.Interval = 4000:RotateRightHead.Interval = 2000
                Case 2:PlaySound "":RotateLeftHead.Interval = 4000:RotateRightHead.Interval = 2000
                Case 3:PlaySound "":RotateLeftHead.Interval = 4000:RotateRightHead.Interval = 2000
                Case 4:PlaySound "":RotateLeftHead.Interval = 4000:RotateRightHead.Interval = 2000
                Case 5:PlaySound "":RotateLeftHead.Interval = 4000:RotateRightHead.Interval = 2000
            End Select
        Else
            RotateLeftHead.Interval = 4000
            RotateRightHead.Interval = 2000
        End If
        RotateRightHead.Enabled = 1
    End If
    If LeftStep> 0 Then
        LeftStep = 0
        RotateLeftHead.Enabled = 0
        cheechspotlight.state = 0
        cheechspotlightglow.state = 0
    End If
    CheechHead.RotY = LeftStep
End Sub

Sub RotateRightHead_Timer
    chongspotlight.state = 1
    chongspotlightglowA.state = 1
    chongspotlightglowB.state = 1
    flasherbirdcage.visible = 1
    RightStep = RightStep + RightDir
    RotateRightHead.Interval = 10
    If RightStep> 40 Then
        RightStep = 40
        RightDir = -1
        If SaySomething Then
            Select Case WhatToSay
                Case 0:PlaySound "sound1chong":RotateRightHead.Interval = 2000
                Case 1:PlaySound "sound2chong":RotateRightHead.Interval = 2000
                Case 2:PlaySound "sound3chong":RotateRightHead.Interval = 2000
                Case 3:PlaySound "sound4chong":RotateRightHead.Interval = 2000
                Case 4:PlaySound "sound5chong":RotateRightHead.Interval = 2000
                Case 5:PlaySound "sound6chong":RotateRightHead.Interval = 2000
            End Select
        Else
            RotateRightHead.Interval = 2000
        End If
    End If
    If RightStep <0 Then
        RightStep = 0
        RotateRightHead.Enabled = 0
        chongspotlight.state = 0
        chongspotlightglowA.state = 0
        chongspotlightglowB.state = 0
        flasherbirdcage.visible = 0
    End If
    ChongHead.RotY = RightStep
End Sub

Sub kicker2_Hit
    PlaySoundAtVol "fx_hole_enter", ActiveBall, 1
    scooplightright.state = 0
    If bsBO.Balls Then
        kicker2.kick 0, 16
    Else
        kicker2.Destroyball
        sw28.CreateBall
        bsBO.AddBall 0
    End If
End Sub

'Lower gates sound and head rotations

Dim GateSoundActivated
GateSoundActivated = False

Sub Gate1a_Hit:GateSoundActivated = True:End Sub

Sub Gate1b_Hit
    If GateSoundActivated Then
Dim w
  w = INT(4 * RND(1) )
  Select Case w
  Case 0:PlaySound "fx_saveyourballs", 0, 1, 0.1, 0.05
        GateSoundActivated = False
        ' rotate Chong
        SaySomething = True
        RightDir = -.75
        RightStep = 0
        RotateChong.Interval = 10
        RotateChong.Enabled = 1
    Case 1:PlaySound "fx_saveyourballs", 0, 1, 0.1, 0.05
        GateSoundActivated = False
        ' rotate Chong
        SaySomething = True
        RightDir = -.75
        RightStep = 0
        RotateChong.Interval = 10
        RotateChong.Enabled = 1
    Case 2:PlaySoundAtVol "fx_bottomgate", Gate1b, 1
    Case 3:PlaySoundAtVol "fx_bottomgate", Gate1b, 1
End Select
End If
End Sub

Sub RotateChong_Timer
    chongspotlight.state = 1
    chongspotlightglowA.state = 1
    chongspotlightglowB.state = 1
    flasherbirdcage.visible = 1
    RightStep = RightStep + RightDir
    RotateChong.Interval = 10
    If RightStep <-15 Then
        RightStep = -15
        RightDir = 1
        If SaySomething Then
            WhatToSay = INT(RND(1) * 4)
            Select Case WhatToSay
                Case 0:PlaySound "sound5chong":RotateChong.Interval = 2000
                Case 1:PlaySound "sound6chong":RotateChong.Interval = 2000
                Case 2:PlaySound "sound7chong":RotateChong.Interval = 2000
                Case 3:PlaySound "sound8chong":RotateChong.Interval = 2000
            End Select
        End If
    End If
    If RightStep> 0 Then
        RightStep = 0
        RotateChong.Enabled = 0
        chongspotlight.state = 0
        chongspotlightglowA.state = 0
        chongspotlightglowB.state = 0
        flasherbirdcage.visible = 0
    End If
    ChongHead.RotY = RightStep
End Sub

Sub Gate3a_Hit:GateSoundActivated = True:End Sub

Sub Gate3b_Hit
If GateSoundActivated Then
Dim w
  w = INT(4 * RND(1) )
  Select Case w
  Case 0:PlaySound "fx_nicebankchot", 0, 1, -0.1, 0.05
        GateSoundActivated = False
        ' rotate Cheech
        SaySomething = True
        LeftDir = .75
        LeftStep = 0
        RotateCheech.Interval = 10
        RotateCheech.Enabled = 1
    Case 1:PlaySound "fx_nicebankchot", 0, 1, -0.1, 0.05
        GateSoundActivated = False
        ' rotate Cheech
        SaySomething = True
        LeftDir = .75
        LeftStep = 0
        RotateCheech.Interval = 10
        RotateCheech.Enabled = 1
    Case 2:PlaySoundAtVol "fx_bottomgate", Gate3b, 1
    Case 3:PlaySoundAtVol "fx_bottomgate", Gate3b, 1
End Select
End If
End Sub

Sub RotateCheech_Timer
    cheechspotlight.state = 1
    cheechspotlightglow.state = 1
    LeftStep = LeftStep + LeftDir
    RotateCheech.Interval = 10
    If LeftStep> 15 Then
        LeftStep = 15
        LeftDir = -1
        If SaySomething Then
            WhatToSay = INT(RND(1) * 4)
            Select Case WhatToSay
                Case 0:PlaySound "sound5cheech":RotateCheech.Interval = 2500
                Case 1:PlaySound "sound6cheech":RotateCheech.Interval = 2500
                Case 2:PlaySound "sound7cheech":RotateCheech.Interval = 2500
                Case 3:PlaySound "sound8cheech":RotateCheech.Interval = 2500
            End Select
        End If
    End If
    If LeftStep <0 Then
        LeftStep = 0
        RotateCheech.Enabled = 0
        cheechspotlight.state = 0
        cheechspotlightglow.state = 0
    End If
    CheechHead.RotY = LeftStep
End Sub

' After a drain, ensure the heads are all in the rest position
Sub ReturnHeads
    LeftDir = 0
    RightDir = 0
    If CheechHead.RotY > 0 Then LeftDir = -1
    If CheechHead.RotY < 0 Then LeftDir = 1
    If ChongHead.RotY > 0 Then RightDir = -1
    If ChongHead.RotY < 0 Then RightDir = 1
    Resetheads.Interval = 10
    Resetheads.Enabled = 1
    cheechspotlight.state = 0
    cheechspotlightglow.state = 0
    chongspotlight.state = 0
    chongspotlightglowA.state = 0
    chongspotlightglowB.state = 0
    flasherbirdcage.visible = 0
End Sub

Sub ResetHeads_Timer
    If CheechHead.RotY = 0 AND ChongHead.RotY = 0 Then
        ResetHeads.Enabled = 0
    End If
    If CheechHead.RotY <> 0 Then CheechHead.RotY = CheechHead.RotY + LeftDir
    If ChongHead.RotY <> 0 Then ChongHead.RotY = ChongHead.RotY + RightDir
End Sub

Sub startgametimer_timer '****make random sound for starting game
    Dim v
    if l0.state = 1 then
        v = INT(6 * RND(1))
        Select Case v
            Case 0:PlaySound "fx_davesnothere"
            Case 1:PlaySound ""
            Case 2:PlaySound ""
            Case 3:PlaySound "fx_letspartyman"
            Case 4:PlaySound ""
            Case 5:PlaySound ""
End Select
    End If
    me.enabled = 0
End Sub

' Rollovers
Sub sw31_Hit:Controller.Switch(31) = 1:PlaySoundAtVol "fx_sensor", ActiveBall, 1:End Sub
Sub sw31_UnHit:Controller.Switch(31) = 0:End Sub

Sub sw32_Hit:Controller.Switch(32) = 1:PlaySoundAtVol "fx_sensor", ActiveBall, 1:End Sub
Sub sw32_UnHit:Controller.Switch(32) = 0:End Sub

Sub sw33_Hit:Controller.Switch(33) = 1:PlaySoundAtVol "fx_sensor", ActiveBall, 1:End Sub
Sub sw33_UnHit:Controller.Switch(33) = 0:End Sub

Sub sw34_Hit:Controller.Switch(34) = 1:PlaySoundAtVol "fx_sensor", ActiveBall, 1:End Sub
Sub sw34_UnHit:Controller.Switch(34) = 0:End Sub

Sub sw74_Hit:Controller.Switch(74) = 1:PlaySoundAtVol "fx_sensor", ActiveBall, 1:End Sub
Sub sw74_UnHit:Controller.Switch(74) = 0:End Sub

Sub sw73_Hit:Controller.Switch(73) = 1:PlaySoundAtVol "fx_sensor", ActiveBall, 1:End Sub
Sub sw73_UnHit:Controller.Switch(73) = 0:End Sub

Sub sw41_Hit:Controller.Switch(41) = 1:PlaySoundAtVol "fx_sensor", ActiveBall, 1:End Sub
Sub sw41_UnHit:Controller.Switch(41) = 0:End Sub

Sub sw42_Hit:Controller.Switch(42) = 1:PlaySoundAtVol "fx_sensor", ActiveBall, 1:End Sub
Sub sw42_UnHit:Controller.Switch(42) = 0:End Sub

Sub sw43_Hit:Controller.Switch(43) = 1:PlaySoundAtVol "fx_sensor", ActiveBall, 1:End Sub
Sub sw43_UnHit:Controller.Switch(43) = 0:End Sub

Sub sw72_Hit:Controller.Switch(72) = 1:PlaySoundAtVol "fx_sensor", ActiveBall, 1:End Sub
Sub sw72_UnHit:Controller.Switch(72) = 0:End Sub

Sub sw71_Hit:Controller.Switch(71) = 1:PlaySoundAtVol "fx_sensor", ActiveBall, 1:End Sub
Sub sw71_UnHit:Controller.Switch(71) = 0:End Sub

Sub sw75_Hit:Controller.Switch(75) = 1:PlaySoundAtVol "fx_sensor", ActiveBall, 1:
Dim bulb
    For each bulb in aGIlights
       bulb.state = 0
   Next
End Sub
Sub sw75_UnHit:Controller.Switch(75) = 0:End Sub

Sub swLocationAdvanceLeft
if Activeball.vely > 0 then
    vpmTimer.PulseSw 68
    PlaySound ""
End if
End Sub

Sub swLocationAdvanceRight
if Activeball.vely > 0 then
    vpmTimer.PulseSw 67
    PlaySound ""
End if
End Sub

' Spinners
Sub Spinner1_Spin
PlaySoundAtVol "fx_spinner", Spinner1, 1
End Sub

Sub Spinner2_Spin
PlaySoundAtVol "fx_spinner", Spinner2, 1
End Sub

'Targets
Sub sw57_Hit:vpmTimer.PulseSw 57:PlaySoundAtVol SoundFX("fx_target", DOFDropTargets), ActiveBall, 1:End Sub
Sub sw58_Hit:vpmTimer.PulseSw 58:PlaySoundAtVol SoundFX("fx_target", DOFDropTargets), ActiveBall, 1:End Sub
Dim sw61Hits:sw61Hits = 0
Sub sw61_Hit:vpmTimer.PulseSw 61:PlaySoundAtVol SoundFX("fx_target", DOFDropTargets), ActiveBall, 1:
    sw61Hits = sw61Hits + 1
    If sw61Hits = 1 Then
    me.uservalue=1
    me.timerenabled= 1
Dim w
  w = INT(7 * RND(1) )
  Select Case w
  Case 0:PlaySound"fx_thatsheavy":CheechSpotlight.duration 1, 1500, 0:CheechSpotlightglow.duration 1, 1500, 0
    Case 1:PlaySound"fx_wowman":CheechSpotlight.duration 1, 1350, 0:CheechSpotlightglow.duration 1, 1350, 0
    Case 2:PlaySound"fx_cheechlaugh":CheechSpotlight.duration 1, 1400, 0:CheechSpotlightglow.duration 1, 1400, 0
    Case 3:PlaySound""
    Case 4:PlaySound""
    Case 5:PlaySound""
    Case 6:PlaySound""
End Select
    sw61Hits = 0
    End If
End Sub
Sub sw62_Hit:vpmTimer.PulseSw 62:PlaySoundAtVol SoundFX("fx_target", DOFDropTargets), ActiveBall, 1:End Sub
Sub sw63_Hit:vpmTimer.PulseSw 63:PlaySoundAtVol SoundFX("fx_target", DOFDropTargets), ActiveBall, 1:End Sub

'Sounds of Characters
Dim sw64Hits:sw64Hits = 0
Dim sw65Hits:sw65Hits = 0
Dim swCousinRedHits:swCousinRedHits = 0
Dim swSargeantStedankoHits:swSargeantStedankoHits = 0

Sub sw64_Hit
    vpmTimer.PulseSw 64
    PlaySoundAtVol SoundFX("fx_target", DOFDropTargets), ActiveBall, 1
    sw64Hits = sw64Hits + 1
    If sw64Hits = 6 Then
    me.uservalue=1
    me.timerenabled= 1
Dim w
  w = INT(2 * RND(1) )
  Select Case w
  Case 0:PlaySound"fx_wherehaveyoubeen":AjaxGirlLight.duration 2, 1900, 0
  Case 1:PlaySound"fx_girlduck":AjaxGirlLight.duration 2, 2500, 0
End Select
    sw64Hits = 0
    End If
End Sub

Sub sw65_Hit
    vpmTimer.PulseSw 65
    PlaySoundAtVol SoundFX("fx_target", DOFDropTargets), ActiveBall, 1
    sw65Hits = sw65Hits + 1
    If sw65Hits = 6 Then
    me.uservalue=1
    me.timerenabled= 1
Dim w
  w = INT(2 * RND(1) )
  Select Case w
  Case 0:PlaySound"fx_whatyoulookinat":StrawberryLight.duration 2, 1200, 0
  Case 1:PlaySound"fx_greatoutdoors":StrawberryLight.duration 2, 2900, 0
End Select
    sw65Hits = 0
    End If
End Sub

Sub sw66_Hit
    vpmTimer.PulseSw 66
    PlaySoundAtVol SoundFX("fx_target", DOFDropTargets), ActiveBall, 1
    sw66Hits = sw66Hits + 1
    If sw66Hits = 10 Then
        StartBirdAnimations
        sw66Hits = 0
        Playsound "fx_itsabirdman"
    End If
End Sub

Sub swCousinRed_Hit
    swCousinRedHits = swCousinRedHits + 1
    If swCousinRedHits = 6 Then
    me.uservalue=1
    me.timerenabled= 1
Dim w
  w = INT(2 * RND(1) )
  Select Case w
  Case 0:PlaySound"fx_callmered":CousinRedLight.duration 2, 1550, 0
  Case 1:PlaySound"fx_holysheepshit":CousinRedLight.duration 2, 3000, 0
End Select
    swCousinRedHits = 0
    End If
End Sub

Sub swSargeantStedanko_Hit
    swSargeantStedankoHits = swSargeantStedankoHits + 1
    If swSargeantStedankoHits = 3 Then
    me.uservalue=1
    me.timerenabled= 1
Dim w
  w = INT(2 * RND(1) )
  Select Case w
  Case 0:PlaySound"fx_shootthemoon":SargeantStedankoLight.duration 2, 1400, 0
  Case 1:PlaySound"fx_sargeantstedanko":SargeantStedankoLight.duration 2, 1500, 0
End Select
    swSargeantStedankoHits = 0
    End If
End Sub

' rotations of Chong and the bird at sw66
Dim sw66Hits:sw66Hits = 0
Dim ChongDir2, ChongStep2
Dim BirdDir, BirdStep

Sub StartBirdAnimations
    ChongDir2 = -1
    ChongStep2 = 0
    BirdDir = 2
    BirdStep = 0
    RotateChong2.Interval = 10
    RotateChong2.Enabled = 1
End Sub

Sub RotateChong2_Timer
    chongspotlight.state = 1
    chongspotlightglowA.state = 1
    chongspotlightglowB.state = 1
    flasherbirdcage.visible = 1
    ChongStep2 = ChongStep2 + ChongDir2
    RotateChong2.Interval = 10
    If ChongStep2 <-40 Then
        ChongStep2 = -40
        ChongDir2 = 1
        PlaySound "ChongSaysSomething"
        RotateChong2.Interval = 6000
        RotateBird.Interval = 3000
        RotateBird.Enabled = 1
    End If
    If ChongStep2> 0 Then
        ChongStep2 = 0
        RotateChong2.Enabled = 0
        chongspotlight.state = 0
        chongspotlightglowA.state = 0
        chongspotlightglowB.state = 0
        flasherbirdcage.visible = 0
    End If
    ChongHead.RotY = ChongStep2
End Sub

Sub RotateBird_Timer
    birdlight.State = 1
    BirdStep = BirdStep + BirdDir
    RotateBird.Interval = 10
    If BirdStep> 118 Then
        BirdStep = 118
        BirdDir = -2
        PlaySound "BirdSaysSomething"
        RotateBird.Interval = 2500
    End If
    If BirdStep <0 Then
        BirdStep = 0
        RotateBird.Enabled = 0
        birdlight.State = 0
    End If
    Bird.RotY = BirdStep
End Sub

'Droptargets (only sound effect)

Sub sw51_Hit:PlaySoundAtVol SoundFX("fx_droptarget", DOFDropTargets), ActiveBall, 1:End Sub

Dim sw52Hits:sw52Hits = 0
Sub sw52_Hit:PlaySoundAtVol SoundFX("fx_droptarget", DOFDropTargets), ActiveBall, 1:
    sw52Hits = sw52Hits + 1
    If sw52Hits = 1 Then
    me.uservalue=1
    me.timerenabled= 1
Dim w
  w = INT(7 * RND(1) )
  Select Case w
  Case 0:PlaySound"fx_imtrippin":ChongSpotlight.duration 1, 1650, 0:ChongSpotlightglowA.duration 1, 1650, 0:ChongSpotlightglowB.duration 1, 1650, 0
    Case 1:PlaySound"fx_chonglaughA":ChongSpotlight.duration 1, 1000, 0:ChongSpotlightglowA.duration 1, 1000, 0:ChongSpotlightglowB.duration 1, 1000, 0
    Case 2:PlaySound"fx_chonglaughB":ChongSpotlight.duration 1, 1200, 0:ChongSpotlightglowA.duration 1, 1200, 0:ChongSpotlightglowB.duration 1, 1200, 0
    Case 3:PlaySound""
    Case 4:PlaySound""
    Case 5:PlaySound""
    Case 6:PlaySound""
End Select
    sw52Hits = 0
    End If
End Sub

Sub sw53_Hit:PlaySoundAtVol SoundFX("fx_droptarget", DOFDropTargets), ActiveBall, 1:End Sub

Sub sw54_Hit:
PlaySoundAtVol SoundFX("fx_droptarget", DOFDropTargets), ActiveBall, 1
:Playsound "fx_popcorn":projector1a.duration 1, 5000, 0:projector1b.duration 1, 5000, 0
:projector2int.duration 1, 5000, 0:gi59.duration 1, 5000, 0:
me.uservalue=1
me.timerenabled= 1
Dim w
  w = INT(6 * RND(1) )
  Select Case w
  Case 0:PlaySound"fx_snackbarclosed"
  Case 1:PlaySound""
    Case 2:PlaySound""
    Case 3:PlaySound""
    Case 4:PlaySound""
    Case 5:PlaySound""
End Select
End Sub

Sub sw55_Hit:
PlaySoundAtVol SoundFX("fx_droptarget", DOFDropTargets), ActiveBall, 1
:Playsound "fx_slurp":projector1a.duration 1, 5000, 0:projector1b.duration 1, 5000, 0
:projector2bug.duration 1, 5000, 0:gi59.duration 1, 5000, 0:
me.uservalue=1
me.timerenabled= 1
Dim w
  w = INT(6 * RND(1) )
  Select Case w
  Case 0:PlaySound"fx_buggery"
  Case 1:PlaySound""
    Case 2:PlaySound""
    Case 3:PlaySound""
    Case 4:PlaySound""
    Case 5:PlaySound""
End Select
End Sub

Sub sw56_Hit:
PlaySoundAtVol SoundFX("fx_droptarget", DOFDropTargets), ActiveBall, 1
:Playsound "fx_bite":projector1a.duration 1, 5000, 0:projector1b.duration 1, 5000, 0
:projector2cha.duration 1, 5000, 0:gi59.duration 1, 5000, 0:
me.uservalue=1
me.timerenabled= 1
Dim w
  w = INT(7 * RND(1) )
  Select Case w
  Case 0:PlaySound"fx_mrtaylor"
  Case 1:PlaySound""
    Case 2:PlaySound""
    Case 3:PlaySound""
    Case 4:PlaySound""
    Case 5:PlaySound""
    Case 6:PlaySound""
End Select
End Sub

'*********
'Solenoids
'*********

SolCallback(1) = "bsTrough.SolIn"
SolCallback(2) = "bsTrough.SolOut"
SolCallback(3) = "dtL.SolDropUp"
SolCallback(4) = "dtM.SolDropUp"
SolCallback(7) = "vpmSolSound""fx_knocker"","
SolCallback(9) = "bsBO.SolOut"
SolCallback(10) = "bsRI.SolOut"
SolCallback(11) = "bsLE.SolOut"
SolCallback(12) = "RGate.Open ="
SolCallback(13) = "LGate.Open ="

'Flashers
SolCallback(17) = "SetLamp 117,"
SolCallback(18) = "SetLamp 118,"
SolCallback(19) = "SetLamp 119,"
SolCallback(20) = "SetLamp 120,"
SolCallback(21) = "SetLamp 121,"
SolCallback(22) = "SetLamp 122,"
SolCallback(23) = "SetLamp 123,"
SolCallback(24) = "SetLamp 124,"
SolCallback(25) = "SetLamp 125,"
SolCallback(26) = "SetLamp 126,"
SolCallback(28) = "SetLamp 128,"


'****************************
'        FLASHERS
'****************************

Sub FlasherTimer_Timer
    if f20.state = 1 then
        GreenLightlitLEFT.visible = 1
    else
        GreenLightlitLEFT.visible = 0
    end If

    if f19.state = 1 then
        GreenLightlitRIGHT.visible = 1
    else
        GreenLightlitRIGHT.visible = 0
    end If

    if li85.state = 1 then
        PedroSpotlight.visible = 1
        ManSpotlight.visible = 1
        RightSpotlight.visible = 1
        PedroSpotlightLight.state = 1
        ManSpotlightLight.state = 1
        RightSpotlightLight.state = 1
    else
        PedroSpotlight.visible = 0
        ManSpotlight.visible = 0
        RightSpotlight.visible = 0
        PedroSpotlightLight.state = 0
        ManSpotlightLight.state = 0
        RightSpotlightLight.state = 0
    end If

    if projector2int.state = 1 then
        intermission.visible = 1
        snackbardoorsint.visible = 1
    else
        intermission.visible = 0
        snackbardoorsint.visible = 0
    end If

    if projector2bug.state = 1 then
        buggery.visible = 1
        snackbardoorsbug.visible = 1
    else
        buggery.visible = 0
        snackbardoorsbug.visible = 0
    end If

    if projector2cha.state = 1 then
        charliechan.visible = 1
        snackbardoorscha.visible = 1
    else
        charliechan.visible = 0
        snackbardoorscha.visible = 0
    end If

    if roadclosed.isdropped then
        roadclosedflash.visible = 0
        roadclosedlight.state = 0
    else
        roadclosedflash.visible = 1
        roadclosedlight.state = 1
    end If

    if troughlight.state = 1 then
        apronplungerglow.state = 1
        plungerglow.state = 1
        ballguardglow.visible = 1
    else
        apronplungerglow.state = 0
        plungerglow.state = 0
        ballguardglow.visible = 0
    end If

    if factorylight.state = 1 then
        factorywindows.visible = 1
        gi24.state = 1
        factorywindowslight.state = 1
        factorywindowslightglow.state = 1
        factorybackwallglow.visible = 1
        factorywallglow.visible = 1
        watertowerglow.visible = 1
    else
        factorywindows.visible = 0
        gi24.state = 0
        factorywindowslight.state = 0
        factorywindowslightglow.state = 0
        factorybackwallglow.visible = 0
        factorywallglow.visible = 0
        watertowerglow.visible = 0
    end If

    if li33.state = 1 then
        CuldesacLight.state = 1
    else
        CuldesacLight.state = 0
    end If

    if culdesac1light.state = 1 then
        culdesac1glow.state = 1
        culdesac1flash.visible = 1
    else
        culdesac1glow.state = 0
        culdesac1flash.visible = 0
    end If

    if culdesac2light.state = 1 then
        culdesac2glow.state = 1
    else
        culdesac2glow.state = 0
    end If

    if culdesac3light.state = 1 then
        culdesac3glow.state = 1
    else
        culdesac3glow.state = 0
    end If

    if culdesac4light.state = 1 then
        culdesac4glow.state = 1
        culdesac4flash.visible = 1
    else
        culdesac4glow.state = 0
        culdesac4flash.visible = 0
    end If

    if li63.state = 1 then
        checkpointlight2.state = 1
    else
        checkpointlight2.state = 0
    end If

    if cheechspotlight.state = 1 then
        cheechbody.image = "Cheech Body Lit"
        cheechhead.image = "Cheech Head Lit"
    else
        cheechbody.image = "Cheech Body"
        cheechhead.image = "Cheech Head"
    end If

    if chongspotlight.state = 1 then
        chongbody.image = "Chongbodylit"
        chonghead.image = "Chongheadlit"
        flasherbirdcage.visible = 1
    else
        chongbody.image = "Chong Body"
        chonghead.image = "Chong Head"
        flasherbirdcage.visible = 0
    end If

    if strawhouselightB.state = 1 then
        strawhousechimlight.state = 1
        strawhousechimlightflasher.visible = 1
        strawhouseflasher.visible = 1
        strawberryhousewindows.visible = 1
    else
        strawhousechimlight.state = 0
        strawhousechimlightflasher.visible = 0
        strawhouseflasher.visible = 0
        strawberryhousewindows.visible = 0
    end If

    if billboardlight1.state = 1 then
        flasherbillboard1A.visible = 1
        flasherbillboard1B.visible = 1
        flasherbillboard1bottom.visible = 1
    else
        flasherbillboard1A.visible = 0
        flasherbillboard1B.visible = 0
        flasherbillboard1bottom.visible = 0
    end If

    if billboardlight2.state = 1 then
        flasherbillboard2A.visible = 1
        flasherbillboard2B.visible = 1
        flasherbillboard2bottom.visible = 1
    else
        flasherbillboard2A.visible = 0
        flasherbillboard2B.visible = 0
        flasherbillboard2bottom.visible = 0
    end If

    if billboardlight3.state = 1 then
        flasherbillboard3A.visible = 1
        flasherbillboard3B.visible = 1
        flasherbillboard3bottom.visible = 1
    else
        flasherbillboard3A.visible = 0
        flasherbillboard3B.visible = 0
        flasherbillboard3bottom.visible = 0
    end If

    if birdlight.state = 1 then
        bird.Image = "birdlit"
        flasherinsidebirdcageglowRIGHT.visible = 1
        flasherinsidebirdcageglowLEFT.visible = 1
    else
        bird.Image = "bird"
        flasherinsidebirdcageglowRIGHT.visible = 0
        flasherinsidebirdcageglowLEFT.visible = 0
    end If

    if vanwindowslight.state = 1 then
        vanwindows.visible = 1
        vanventlight.visible = 1
        vanheadandtaillights.visible = 1
    else
        vanwindows.visible = 0
        vanventlight.visible = 0
        vanheadandtaillights.visible = 0
    end If

    if gi57.state = 1 then
        LoveMachineHeadlights.visible = 1
    else
        LoveMachineHeadlights.visible = 0
    end If

    if lightf21.state = 1 then
        f21bulblit.visible = 1
    else
        f21bulblit.visible = 0
    end If

    if scooplightleft.state = 1 then
        ScooplightLEFTlit.visible = 1
    else
        ScooplightLEFTlit.visible = 0
    end If

    if scooplightright.state = 1 then
        ScooplightRIGHTlit.visible = 1
    else
        ScooplightRIGHTlit.visible = 0
    end If

    if gi18.state = 1 then
        targetshadow.visible = 1
    else
        targetshadow.visible = 0
    end If

    if sw56.isdropped then
        sw56DTshadow.visible = 0
    else
        sw56DTshadow.visible = 1
    end If

    if sw55.isdropped then
        sw55DTshadow.visible = 0
    else
        sw55DTshadow.visible = 1
    end If

    if sw54.isdropped then
        sw54DTshadow.visible = 0
    else
        sw54DTshadow.visible = 1
    end If

    if sw53.isdropped then
        sw53DTshadow.visible = 0
    else
        sw53DTshadow.visible = 1
    end If

    if sw52.isdropped then
        sw52DTshadow.visible = 0
    else
        sw52DTshadow.visible = 1
    end If

    if sw51.isdropped then
        sw51DTshadow.visible = 0
    else
        sw51DTshadow.visible = 1
    end If

    if gi12.state = 0 then
        sw54DTshadow.visible = 0
        sw55DTshadow.visible = 0
        sw56DTshadow.visible = 0
    end If

    if gi15.state = 0 then
        sw51DTshadow.visible = 0
        sw52DTshadow.visible = 0
        sw53DTshadow.visible = 0
    end If

    if gi69.state = 0 then
        roadclosedflash.visible = 0
        roadclosedlight.state = 0
    end If

    if gi50.state = 0 then
        scooplightLEFT.state = 0
        strawhousechimlight.state = 0
        strawhouselightA.state = 0
        strawhouselightB.state = 0
        strawhousechimlightflasher.visible = 0
        strawhouseflasher.visible = 0
    end If

If Table1.ShowDT = true then
if li41.state = 1 then
        breflectDT.visible = 1
    else
        breflectDT.visible = 0
    end If

if li44.state = 1 then
        sreflectDT.visible = 1
    else
        sreflectDT.visible = 0
    end If

Else 'if in FS mode


if li41.state = 1 then
        breflectCAB.visible = 1
    else
        breflectCAB.visible = 0
    end If

if li44.state = 1 then
        sreflectCAB.visible = 1
    else
        sreflectCAB.visible = 0
    end If
End If


    if gi57.state = 0 then
        RunnerRedLED.state = 0
        RunnerRedLEDa.state = 0
        RunnerRedLEDb.state = 0
        RunnerRedLEDc.state = 0
        RunnerRedLEDd.state = 0
        RunnerRedLEDe.state = 0
        RunnerBlueLED.state = 0
        RunnerBlueLEDa.state = 0
        RunnerBlueLEDb.state = 0
        RunnerBlueLEDc.state = 0
        RunnerBlueLEDd.state = 0
        RunnerBlueLEDe.state = 0
    end If

if gi41.state = 1 then
        gi41flash.visible = 1
    else
        gi41flash.visible = 0
    end If

if gi64.state = 1 then
        gi64flash.visible = 1
    else
        gi64flash.visible = 0
    end If

if gi5.state = 1 then
        outlaneleftshadow.visible = 0
    else
        outlaneleftshadow.visible = 1
    end If

if gi10.state = 1 then
        outlanerightshadow.visible = 0
    else
        outlanerightshadow.visible = 1
    end If

if boxlight.state = 1 then
        boxlightflasher.visible = 1
        clownjointlit1.visible = 1
    else
        boxlightflasher.visible = 0
        clownjointlit1.visible = 0
    end If

if gi55.state = 0 then
        boxlight.state = 0
    end If

if DrainpostLEDglow.state = 1 then
        DrainpostLEDlit.visible = 1
    else
        DrainpostLEDlit.visible = 0
    end If

End Sub

'**************
' Flipper Subs
'**************

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
    If Enabled Then
        PlaySoundAtVol SoundFX("fx_flipperup", DOFFlippers), LeftFlipper, 1
        LeftFlipper.RotateToEnd
        PlaySoundAtVol SoundFX("fx_flipperup", DOFFlippers), UpperFLipper, 1
        UpperFlipper.RotateToEnd
    Else
        PlaySoundAtVol SoundFX("fx_flipperdown", DOFFlippers), LeftFlipper, 1
        LeftFlipper.RotateToStart
        PlaySoundAtVol SoundFX("fx_flipperdown", DOFFlippers), UpperFLipper, 1
        UpperFlipper.RotateToStart
    End If
End Sub

Sub SolRFlipper(Enabled)
    If Enabled Then
        PlaySoundAtVol SoundFX("fx_flipperup", DOFFlippers), RightFlipper, 1
        RightFlipper.RotateToEnd
    Else
        PlaySoundAtVol SoundFX("fx_flipperdown", DOFFlippers), RightFlipper, 1
        RightFlipper.RotateToStart
    End If
End Sub

Sub LeftFlipper_Collide(parm)
    PlaySoundAtVol "fx_rubber_flipper", LeftFlipper, parm / 10
End Sub

Sub UpperFlipper_Collide(parm)
    PlaySoundAtVol "fx_rubber_flipper", UpperFlipper, parm / 10
End Sub

Sub RightFlipper_Collide(parm)
    PlaySoundAtVol "fx_rubber_flipper", RightFlipper, parm / 10
End Sub


'**********************************************************
'     JP's Flasher Fading for VPX and Vpinmame v 1.0
'       (Based on Pacdude's Fading Light System)
' This is a fast fading for the Flashers in vpinmame tables
'**********************************************************

Dim FadingState(200)

InitLamps() ' turn off the lights and flashers and reset them to the default parameters

' vpinmame Lamp & Flasher Timers

Sub LampTimer_Timer()
    Dim chgLamp, num, chg, ii
    chgLamp = Controller.ChangedLamps
    If Not IsEmpty(chgLamp)Then
        For ii = 0 To UBound(chgLamp)
            FadingState(chgLamp(ii, 0)) = chgLamp(ii, 1) + 3 'fading step
        Next
    End If
    UpdateLamps
    UpdateLeds
End Sub

Sub UpdateLamps
    ' playfield lights
    Lampm 11, li11
    Flash 11, li11a
    Lampm 12, li12
    Flash 12, li12a
    Lamp 13, li13
    Lamp 14, li14
    Lamp 15, li15
    Lamp 16, li16
    Flash 17, li17
    Flash 18, li18
    Flash 21, li21
    Flash 22, li22
    Lamp 23, li23
    Lamp 24, li24
    Lamp 25, li25
    Lamp 26, li26
    Lamp 27, li27
    Lamp 28, li28
    Lamp 31, li31
    Lamp 32, li32
    Lamp 33, li33
    Lamp 34, li34
    Lamp 35, li35
    Lamp 36, li36
    Lamp 37, li37
    Lamp 38, li38
    Lamp 41, li41
    Lamp 42, li42
    Lamp 43, li43
    Lamp 44, li44
    Lamp 45, li45
    Lamp 46, li46
    Lamp 47, li47
    Lamp 48, li48
    Lamp 51, li51
    Lamp 52, li52
    Lamp 53, li53
    Lamp 54, li54
    Lamp 55, li55
    Lamp 56, li56
    Lamp 57, li57
    Lamp 58, li58
    Lamp 61, li61
    Lamp 62, li62
    Lamp 63, li63
    Lamp 64, li64
    Lamp 65, li65
    Lamp 66, li66
    Lamp 67, li67
    Lamp 68, li68
    Lamp 71, li71
    Lamp 72, li72
    Lamp 73, li73
    Lamp 74, li74
    Lamp 75, li75
    Lamp 76, li76
    Lamp 77, li77
    Lamp 78, li78
    Lamp 81, li81
    Lamp 82, li82
    Flash 83, li83
    Flash 84, li84
    Lampm 85, li85
    Flash 85, li85a
    Lamp 86, li86
    Lamp 87, li87
    Lampm 88, li88
    Flash 88, li88a
    'Flashers
    Flashm 117, f17a
    Flash 117, f17
    Flashm 118, f18a
    Flash 118, f18
    Lamp 119, f19
    Lamp 120, f20
    Lampm 121, JadeLight
    Lampm 121, DebbieLight
    Lampm 121, lightf21
    Flashm 121, f21a
    Flash 121, f21
    Lampm 122, f22a
    Lampm 122, f22b
    Flashm 122, flasherJackpotChong
    Flashm 122, flasherJackpotbirdcage
    Flash 122, f22
    Flashm 123, flasherplasticCheech
    Flashm 123, flasherplasticChong
    Lampm 123, Checkpointlight
    Flashm 123, f23a
    Flash 123, f23
    Lampm 124, CheechSpotlight
    Lampm 124, CheechSpotlightglow
    Lampm 124, ChongSpotlight
    Lampm 124, ChongSpotlightglowA
    Lampm 124, ChongSpotlightglowB
    Flashm 124, f24a
    Flash 124, f24
    Lampm 125, lightf25
    Flashm 125, f25a
    Flash 125, f25
    Lampm 126, lightf26
    Flashm 126, f26a
    Flash 126, f26
'   Lamp 128, f28
End Sub

' div lamp subs

Sub InitLamps()
    Dim x
    LampTimer.Interval = 25 ' flasher fading speed
    LampTimer.Enabled = 1
    For x = 0 to 200
        FadingState(x) = 3 ' used to track the fading state
    Next
End Sub

Sub SetLamp(nr, value) ' 0 is off, 1 is on
    FadingState(nr) = abs(value) + 3
End Sub

' Lights: used for VPX standard lights, the fading is handled by VPX itself, they are here to be able to make them work together with the flashers

Sub Lamp(nr, object)
    Select Case FadingState(nr)
        Case 4:object.state = 1:FadingState(nr) = 0
        Case 3:object.state = 0:FadingState(nr) = 0
    End Select
End Sub

Sub Lampm(nr, object) ' used for multiple lights, it doesn't change the fading state
    Select Case FadingState(nr)
        Case 4:object.state = 1
        Case 3:object.state = 0
    End Select
End Sub

' Flashers: 4 is on,3,2,1 fade steps. 0 is off

Sub Flash(nr, object)
    Select Case FadingState(nr)
        Case 4:Object.IntensityScale = 1:FadingState(nr) = 0
        Case 3:Object.IntensityScale = 0.66:FadingState(nr) = 2
        Case 2:Object.IntensityScale = 0.33:FadingState(nr) = 1
        Case 1:Object.IntensityScale = 0:FadingState(nr) = 0
    End Select
End Sub

Sub Flashm(nr, object) 'multiple flashers, it doesn't change the fading state
    Select Case FadingState(nr)
        Case 4:Object.IntensityScale = 1
        Case 3:Object.IntensityScale = 0.66
        Case 2:Object.IntensityScale = 0.33
        Case 1:Object.IntensityScale = 0
    End Select
End Sub

'**********************************************************************************************************
'Digital Display
'**********************************************************************************************************
Dim Digits(32)
Digits(0) = Array(a00, a05, a0c, a0d, a08, a01, a06, a0f, a02, a03, a04, a07, a0b, a0a, a09, a0e)
Digits(1) = Array(a10, a15, a1c, a1d, a18, a11, a16, a1f, a12, a13, a14, a17, a1b, a1a, a19, a1e)
Digits(2) = Array(a20, a25, a2c, a2d, a28, a21, a26, a2f, a22, a23, a24, a27, a2b, a2a, a29, a2e)
Digits(3) = Array(a30, a35, a3c, a3d, a38, a31, a36, a3f, a32, a33, a34, a37, a3b, a3a, a39, a3e)
Digits(4) = Array(a40, a45, a4c, a4d, a48, a41, a46, a4f, a42, a43, a44, a47, a4b, a4a, a49, a4e)
Digits(5) = Array(a50, a55, a5c, a5d, a58, a51, a56, a5f, a52, a53, a54, a57, a5b, a5a, a59, a5e)
Digits(6) = Array(a60, a65, a6c, a6d, a68, a61, a66, a6f, a62, a63, a64, a67, a6b, a6a, a69, a6e)
Digits(7) = Array(a70, a75, a7c, a7d, a78, a71, a76, a7f, a72, a73, a74, a77, a7b, a7a, a79, a7e)
Digits(8) = Array(a80, a85, a8c, a8d, a88, a81, a86, a8f, a82, a83, a84, a87, a8b, a8a, a89, a8e)
Digits(9) = Array(a90, a95, a9c, a9d, a98, a91, a96, a9f, a92, a93, a94, a97, a9b, a9a, a99, a9e)
Digits(10) = Array(aa0, aa5, aac, aad, aa8, aa1, aa6, aaf, aa2, aa3, aa4, aa7, aab, aaa, aa9, aae)
Digits(11) = Array(ab0, ab5, abc, abd, ab8, ab1, ab6, abf, ab2, ab3, ab4, ab7, abb, aba, ab9, abe)
Digits(12) = Array(ac0, ac5, acc, acd, ac8, ac1, ac6, acf, ac2, ac3, ac4, ac7, acb, aca, ac9, ace)
Digits(13) = Array(ad0, ad5, adc, add, ad8, ad1, ad6, adf, ad2, ad3, ad4, ad7, adb, ada, ad9, ade)
Digits(14) = Array(ae0, ae5, aec, aed, ae8, ae1, ae6, aef, ae2, ae3, ae4, ae7, aeb, aea, ae9, aee)
Digits(15) = Array(af0, af5, afc, afd, af8, af1, af6, aff, af2, af3, af4, af7, afb, afa, af9, afe)

Digits(16) = Array(b00, b05, b0c, b0d, b08, b01, b06, b0f, b02, b03, b04, b07, b0b, b0a, b09, b0e)
Digits(17) = Array(b10, b15, b1c, b1d, b18, b11, b16, b1f, b12, b13, b14, b17, b1b, b1a, b19, b1e)
Digits(18) = Array(b20, b25, b2c, b2d, b28, b21, b26, b2f, b22, b23, b24, b27, b2b, b2a, b29, b2e)
Digits(19) = Array(b30, b35, b3c, b3d, b38, b31, b36, b3f, b32, b33, b34, b37, b3b, b3a, b39, b3e)
Digits(20) = Array(b40, b45, b4c, b4d, b48, b41, b46, b4f, b42, b43, b44, b47, b4b, b4a, b49, b4e)
Digits(21) = Array(b50, b55, b5c, b5d, b58, b51, b56, b5f, b52, b53, b54, b57, b5b, b5a, b59, b5e)
Digits(22) = Array(b60, b65, b6c, b6d, b68, b61, b66, b6f, b62, b63, b64, b67, b6b, b6a, b69, b6e)
Digits(23) = Array(b70, b75, b7c, b7d, b78, b71, b76, b7f, b72, b73, b74, b77, b7b, b7a, b79, b7e)
Digits(24) = Array(b80, b85, b8c, b8d, b88, b81, b86, b8f, b82, b83, b84, b87, b8b, b8a, b89, b8e)
Digits(25) = Array(b90, b95, b9c, b9d, b98, b91, b96, b9f, b92, b93, b94, b97, b9b, b9a, b99, b9e)
Digits(26) = Array(ba0, ba5, bac, bad, ba8, ba1, ba6, baf, ba2, ba3, ba4, ba7, bab, baa, ba9, bae)
Digits(27) = Array(bb0, bb5, bbc, bbd, bb8, bb1, bb6, bbf, bb2, bb3, bb4, bb7, bbb, bba, bb9, bbe)
Digits(28) = Array(bc0, bc5, bcc, bcd, bc8, bc1, bc6, bcf, bc2, bc3, bc4, bc7, bcb, bca, bc9, bce)
Digits(29) = Array(bd0, bd5, bdc, bdd, bd8, bd1, bd6, bdf, bd2, bd3, bd4, bd7, bdb, bda, bd9, bde)
Digits(30) = Array(be0, be5, bec, bed, be8, be1, be6, bef, be2, be3, be4, be7, beb, bea, be9, bee)
Digits(31) = Array(bf0, bf5, bfc, bfd, bf8, bf1, bf6, bff, bf2, bf3, bf4, bf7, bfb, bfa, bf9, bfe)

Sub UpdateLeds
    Dim ChgLED, ii, jj, num, chg, stat, obj, b, x
    ChgLED = Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
    If Not IsEmpty(ChgLED)Then
        For ii = 0 To UBound(chgLED)
            num = chgLED(ii, 0):chg = chgLED(ii, 1):stat = chgLED(ii, 2)
            if(num <32)then
                For Each obj In Digits(num)
                    If chg And 1 Then obj.State = stat And 1
                    chg = chg \ 2:stat = stat \ 2
                Next
            Else
            end if
        Next
    End If
End Sub

'******************************
' Diverse Collection Hit Sounds
'******************************

Sub aMetals_Hit(idx):PlaySound "fx_MetalHit2", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aRubber_Bands_Hit(idx):PlaySound "fx_rubber_band", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aRubber_Posts_Hit(idx):PlaySound "fx_rubber_post", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aRubber_Pins_Hit(idx):PlaySound "fx_rubber_pin", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aPlastics_Hit(idx):PlaySound "fx_PlasticHit", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aGates_Hit(idx):PlaySound "fx_Gate", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aWoods_Hit(idx):PlaySound "fx_Woodhit", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub

Sub REnd1_Hit:PlaySound "", 0, 1, 0.05, 0.05:End Sub

' *******************************************************************************************************
' Positional Sound Playback Functions by DJRobX, Rothbauerw, Thalamus and Herweh
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


'*****************************************
'      JP's VP10 Rolling Sounds
'*****************************************

Const tnob = 5 ' total number of balls
Const lob = 0  'number of locked balls
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
    If UBound(BOT) = lob - 1 Then Exit Sub 'there no extra balls on this table

    ' play the rolling sound for each ball
    For b = lob to UBound(BOT)
        If BallVel(BOT(b))> 1 Then
            If BOT(b).z <30 Then
                ballpitch = Pitch(BOT(b))
            Else
                ballpitch = Pitch(BOT(b)) + 15000 'increase the pitch on a ramp or elevated surface
            End If
            rolling(b) = True
            PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b)), Pan(BOT(b)), 0, ballpitch, 1, 0
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
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 2000, Pan(ball1), 0, Pitch(ball1), 0, 0
End Sub

'******************
' RealTime Updates
'******************

Set MotorCallback = GetRef("RealTimeUpdates")

Sub RealTimeUpdates
    RollingUpdate
    ClownUpdate
End Sub

'*************************
' GI - needs new vpinmame
'*************************

Set GICallback = GetRef("GIUpdate")

Sub GIUpdate(no, Enabled)
    For each x in aGiLights
        x.State = ABS(Enabled)
    Next
End Sub

'*****************************************
'     FLIPPER SHADOWS
'*****************************************
'Primitive Flipper Code

Sub FlipperTimer_Timer
    FlipperLSh.RotZ = LeftFlipper.currentangle
    LFLogo.RotZ = LeftFlipper.currentangle
    UpperFlipperSh.RotZ = UpperFlipper.currentangle
    UFLogo.RotZ = UpperFlipper.currentangle
    FlipperRSh.RotZ = RightFlipper.currentangle
    RFLogo.RotZ = RightFlipper.currentangle
End Sub

'*****************************************
' ninuzzu's BALL SHADOW
'*****************************************
Dim BallShadow
BallShadow = Array(BallShadow1, BallShadow2, BallShadow3, BallShadow4, BallShadow5)

Sub BallShadowUpdate_timer()
    Dim BOT, b
    BOT = GetBalls
    ' hide shadow of deleted balls
    If UBound(BOT) <(tnob-1)Then
        For b = (UBound(BOT) + 1)to(tnob-1)
            BallShadow(b).visible = 0
        Next
    End If
    ' exit the Sub if no balls on the table
    If UBound(BOT) = -1 Then Exit Sub
    ' render the shadow for each ball
    For b = 0 to UBound(BOT)
        If BOT(b).X <Table1.Width / 2 Then
            BallShadow(b).X = ((BOT(b).X)-(Ballsize / 6) + ((BOT(b).X -(Table1.Width / 2)) / 7)) + 13
        Else
            BallShadow(b).X = ((BOT(b).X) + (Ballsize / 6) + ((BOT(b).X -(Table1.Width / 2)) / 7))- 13
        End If
        ballShadow(b).Y = BOT(b).Y + 10
        If BOT(b).Z> 20 Then
            BallShadow(b).visible = 1
        Else
            BallShadow(b).visible = 0
        End If
    Next
End Sub

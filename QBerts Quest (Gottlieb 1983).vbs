'Q*bert's Quest 2.0 by bord

Option Explicit
Randomize

' Thalamus 2018-07-24
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Thalamus 2018-09-01 : Improved directional sounds

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
Const VolRH     = 1    ' Rubber hits volume.
Const VolPo     = 1    ' Rubber posts volume.
Const VolPi     = 1    ' Rubber pins volume.
Const VolTarg   = 1    ' Targets volume.
Const VolSpin   = 1.5  ' Spinners volume.
Const VolFlip   = 1    ' Flipper volume.


On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

LoadVPM "01500000", "sys80.vbs", 3.10

'**********************************************************
'********       OPTIONS     *******************************
'**********************************************************

Dim BallShadows: Ballshadows=1          '******************set to 1 to turn on Ball shadows
Dim FlipperShadows: FlipperShadows=1  '***********set to 1 to turn on Flipper shadows

Dim DesktopMode: DesktopMode = Table1.ShowDT
dim hiddenvalue
If DesktopMode = True Then 'Show Desktop components
siderailleft.visible=1
siderailright.visible=1
lockdown.visible=1
hiddenvalue=0
Else
siderailleft.visible=0
siderailright.visible=0
lockdown.visible=0
hiddenvalue = 1
End if

SolCallback(1) = "SolLeftDropUp"
'SolCallback(2) = "bsRSaucer.SolOut"
SolCallback(2) = "SolRightEMKicker"
'SolCallback(3) =   coin1
'SolCallback(4) =   coin2
'SolCallback(5) = "bsLSaucer.SolOut"
SolCallback(5) = "SolLeftEMKicker"
SolCallback(6) = "SolRightDropUp"
'SolCallback(7) =   coin3
SolCallback(8) = "vpmSolSound SoundFX(""Knocker"",DOFKnocker),"
SolCallback(9) = "bsTrough.SolOut"
SolCallback(10) = "FastFlips.TiltSol"

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
     If Enabled Then
         PlaySoundAtVol SoundFX("fx_FlipperupL",DOFFlippers), LeftFlipper, VolFlip:LeftFlipper.RotateToEnd:RightFlipper2.RotateToEnd
     Else
         PlaySoundAtVol SoundFX("fx_Flipperdown",DOFFlippers), LeftFlipper, VolFlip:LeftFlipper.RotateToStart:RightFlipper2.RotateToStart
     End If
  End Sub

Sub SolRFlipper(Enabled)
     If Enabled Then
         PlaySoundAtVol SoundFX("fx_FlipperupR",DOFFlippers), RightFlipper, VolFlip:RightFlipper.RotateToEnd:LeftFlipper2.RotateToEnd
     Else
         PlaySoundAtVol SoundFX("fx_Flipperdown",DOFFlippers), RightFlipper, VolFlip:RightFlipper.RotateToStart:LeftFlipper2.RotateToStart
     End If
End Sub

'***** EM kickers animation

Dim LeftEMpos, RightEMPos

Sub SolLeftEMKicker(enabled)
  If enabled Then
    bsLSaucer.ExitSol_On
    LeftEMpos = 0
    Lemk.RotX = 26
    LeftEMTimer.Enabled = 1
  End If
End Sub

Sub LeftEMTimer_Timer
    Select Case LeftEMpos
        Case 1:Lemk.RotX = 14
        Case 2:Lemk.RotX = 2
        Case 3:Lemk.RotX = -10:LeftEMTimer.Enabled = 0
    End Select
    LeftEMpos = LeftEMpos + 1
End Sub

Sub SolRightEMKicker(enabled)
  If enabled Then
    bsRSaucer.ExitSol_On
    RightEMpos = 0
    Remk.RotX = 26
    RightEMTimer.Enabled = 1
  End If
End Sub

Sub RightEMTimer_Timer
    Select Case RightEMpos
        Case 1:Remk.RotX = 14
        Case 2:Remk.RotX = 2
        Case 3:Remk.RotX = -10:RightEMTimer.Enabled = 0
    End Select
    RightEMpos = RightEMpos + 1
End Sub

''*****GI Lights On
'dim xx
'For each xx in GI:xx.State = 1: Next

DisplayTimer.Enabled = true

'Primitive Flipper
Sub FlipperTimer_Timer
    leftflipper_prim.rotz = LeftFlipper.currentangle  '+ 90
    rightflipper_prim.rotz = RightFlipper.currentangle '+ 45
    LeftFlipper2_prim.rotz = LeftFlipper2.currentangle '+ 45
    RightFlipper2_prim.rotz = RightFlipper2.currentangle '+ 45
    FlipperLSh.rotz = LeftFlipper.currentangle  '+ 45
    FlipperRSh.rotz = RightFlipper.currentangle '+ 45
    FlipperLSh1.rotz = LeftFlipper2.currentangle '+ 45
    FlipperRSh1.rotz = RightFlipper2.currentangle '+ 45
    rightgate1_prim.RotX = Gate7.CurrentAngle + 90
    rightgate2_prim.RotX = Gate1.CurrentAngle + 90
    rightgate3_prim.RotX = Gate2.CurrentAngle + 90
    leftgate1_prim.RotX = Gate5.CurrentAngle + 90
    leftgate2_prim.RotX = Gate3.CurrentAngle + 90
End Sub

Dim bsTrough, dtLBank, dtRBank, bsLSaucer, bsRSaucer, bump1, bump2, FastFlips

Const cGameName = "qbquest"
Const UseSolenoids = 1
Const UseLamps = 0
Const UseSync = 0
Const UseGI = 0
Const SSolenoidOn="SolOn"
Const SSolenoidOff="SolOff"
Const SCoin="coin"
bsize = 26
bmass = 1.5


Sub Table1_Init
    vpmInit Me
    On Error Resume Next
        With Controller
        .GameName = cGameName
        If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
        .SplashInfoLine = "Q*Bert's Quest (Gottlieb 1983)" & vbNewLine & "2.0"
        .HandleMechanics=0
        .HandleKeyboard=0
        .ShowDMDOnly=1
        .ShowFrame=0
        .ShowTitle=0
        .hidden = hiddenvalue
        If Err Then MsgBox Err.Description
    End With
    On Error Goto 0
        Controller.SolMask(0)=0
      vpmTimer.AddTimer 2000,"Controller.SolMask(0)=&Hffffffff'" 'ignore all solenoids - then add the timer to renable all the solenoids after 2 seconds
        Controller.Run
    If Err Then MsgBox Err.Description
    On Error Goto 0

    PinMAMETimer.Interval=PinMAMEInterval
    PinMAMETimer.Enabled=1

    vpmNudge.TiltSwitch=57
    vpmNudge.Sensitivity=5
    vpmNudge.TiltObj = Array(Bumper1,Bumper2,LeftslingShot,RightslingShot, leftflipper, rightflipper, leftflipper2, rightflipper2)

    Kicker3.CreateBall
    Kicker3.Kick 180, 1

 Set FastFlips = new cFastFlips
    with FastFlips
        .CallBackL = "SolLflipper"  'Point these to flipper subs
        .CallBackR = "SolRflipper"  '...
       ' .CallBackUL = "SolULflipper"'...(upper flippers, if needed)
       ' .CallBackUR = "SolURflipper"'...
        .TiltObjects = True 'Optional, if True calls vpmnudge.solgameon automatically. IF YOU GET A LINE 1 ERROR, DISABLE THIS! (or setup vpmNudge.TiltObj!)
    '   .DebugOn = False        'Debug, always-on flippers. Call FastFlips.DebugOn True or False in debugger to enable/disable.
    end with

    ' Trough
    Set bsTrough = New cvpmBallStack
    With bsTrough
        .InitSw 0, 67, 0, 0, 0, 0, 0, 0
        .InitKick BallRelease, 90, 5
        .InitExitSnd  SoundFX("ballrelease",DOFContactors), SoundFX("Solenoid",DOFContactors)
        .Balls=1
    End With

    ' Drop targets left
    set dtLBank = new cvpmdroptarget
    With dtLBank
        .initdrop array(sw00, sw10), array(0, 10)
        .initSnd SoundFX("DTDrop",DOFDropTargets),SoundFX("DTReset",DOFContactors)
    End With

    ' Drop targets right
    set dtRBank = new cvpmdroptarget
    With dtRBank
        .initdrop array(sw50, sw60), array(50, 60)
        .initSnd SoundFX("DTDrop",DOFDropTargets),SoundFX("DTReset",DOFContactors)
    End With

    ' Left Saucer
    Set bsLSaucer = New cvpmBallStack
    With bsLSaucer
        .InitSaucer sw54, 54, 0, 30
        .InitExitSnd  SoundFX("fx_kicker_release",DOFContactors), SoundFX("Solenoid",DOFContactors)
        .KickAngleVar = 2
    End With

    ' Right Saucer
    Set bsRSaucer = New cvpmBallStack
    With bsRSaucer
        .InitSaucer sw4, 4, 0, 28
        .InitExitSnd  SoundFX("fx_kicker_release",DOFContactors), SoundFX("Solenoid",DOFContactors)
        .KickAngleVar = 2
    End With

if ballshadows=1 then
        BallShadowUpdate.enabled=1
      else
        BallShadowUpdate.enabled=0
    end if

    if flippershadows=1 then
        FlipperLSh.visible=1
    FlipperLSh1.visible=1
        FlipperRSh.visible=1
        FlipperRSh1.visible=1
      else
        FlipperLSh.visible=0
        FlipperLSh1.visible=0
        FlipperRSh.visible=0
        FlipperRSh1.visible=0
    end if

End Sub

'**********************************************************************************************************
'Plunger code
'**********************************************************************************************************

Sub Table1_KeyDown(ByVal KeyCode)
    If KeyCode = LeftFlipperKey then FastFlips.FlipL True :  FastFlips.FlipUL True
    If KeyCode = RightFlipperKey then FastFlips.FlipR True :  FastFlips.FlipUR True
    If KeyDownHandler(keycode) Then Exit Sub
    If keycode = PlungerKey Then Plunger.Pullback:playsoundAt "plungerpull", Plunger
End Sub

Sub Table1_KeyUp(ByVal KeyCode)
    If KeyCode = LeftFlipperKey then FastFlips.FlipL False :  FastFlips.FlipUL False
    If KeyCode = RightFlipperKey then FastFlips.FlipR False :  FastFlips.FlipUR False
    If KeyUpHandler(keycode) Then Exit Sub
    If keycode = PlungerKey Then Plunger.Fire:PlaySoundAt "plunger", Plunger
End Sub

'Drop Targets
 Sub Sw00_Dropped:dtLBank.Hit 1 : Ldrop1.state=1
End Sub
 Sub Sw10_Dropped:dtLBank.Hit 2 : Ldrop2.state=1
End Sub
 Sub Sw50_Dropped:dtRBank.Hit 1 : Rdrop1.state=1
End Sub
 Sub Sw60_Dropped:dtRBank.Hit 2 : Rdrop2.state=1
End Sub


Sub SolRightDropUp(enabled)
    dim xx
    if enabled then
        dtRBank.SolDropUp enabled
        For each xx in DTRight: xx.state=0:Next
    end if
End Sub

Sub SolLeftDropUp(enabled)
    dim xx
    if enabled then
        dtLBank.SolDropUp enabled
        For each xx in DTLeft: xx.state=0:Next
    end if
End Sub

'Bumpers
Sub Bumper1_Hit : vpmTimer.PulseSw(6) : playsoundAtVol SoundFXDOF("fx_bumper1",101,DOFPulse,DOFContactors), Bumper1, VolBump: End Sub
Sub Bumper2_Hit : vpmTimer.PulseSw(6) : playsoundAtVol SoundFXDOF("fx_bumper2",102,DOFPulse,DOFContactors), Bumper2, VolBump: End Sub

' Rollovers
Sub sw15_Hit:Controller.Switch(15)=1 : PlaySoundAtVol "rollover", sw15, VolRol:End Sub
Sub sw15_unHit:Controller.Switch(15)=0:End Sub

Sub sw65_Hit:Controller.Switch(65)=1 : PlaySoundAtVol "rollover", sw65, VolRol:End Sub
Sub sw65_unHit:Controller.Switch(65)=0:End Sub

Sub sw55_Hit:Controller.Switch(55)=1 : PlaySoundAtVol "rollover", sw55, VolRol:End Sub
Sub sw55_unHit:Controller.Switch(55)=0:End Sub

Sub sw13_Hit:Controller.Switch(13)=1 : PlaySoundAtVol "rollover", sw13, VolRol:End Sub
Sub sw13_unHit:Controller.Switch(13)=0:End Sub

Sub sw14_Hit:Controller.Switch(14)=1 :: PlaySoundAtVol "rollover", sw14, VolRol:End Sub
Sub sw14_unHit:Controller.Switch(14)=0:End Sub

Sub sw3_Hit:Controller.Switch(3)=1 : PlaySoundAtVol "rollover", sw3, VolRol:End Sub
Sub sw3_unHit:Controller.Switch(3)=0:End Sub

'gates rollovers
Sub sw61_Hit : Controller.Switch(61)=1 : End Sub
Sub sw61_unHit : Controller.Switch(61)=0 : End Sub
Sub sw62_Hit : Controller.Switch(62)=1 : End Sub
Sub sw62_unHit : Controller.Switch(62)=0 : End Sub
Sub sw63_Hit : Controller.Switch(63)=1 : End Sub
Sub sw63_unHit : Controller.Switch(63)=0 : End Sub
Sub sw64_Hit : Controller.Switch(64)=1 : End Sub
Sub sw64_unHit : Controller.Switch(64)=0 : End Sub
Sub sw53_Hit : Controller.Switch(53)=1 : End Sub
Sub sw53_unHit : Controller.Switch(53)=0 : End Sub

'Standup target
Sub sw1_Hit : vpmTimer.PulseSw(1):PlaySoundAtVol SoundFX("target",DOFTargets), sw1, VolTarg:End Sub
Sub sw2_Hit : vpmTimer.PulseSw(2):PlaySoundAtVol SoundFX("target",DOFTargets), sw1, VolTarg:End Sub
Sub sw5_Hit : vpmTimer.PulseSw(5):PlaySoundAtVol SoundFX("target",DOFTargets), sw1, VolTarg:End Sub
Sub sw11_Hit : vpmTimer.PulseSw(11):PlaySoundAtVol SoundFX("target",DOFTargets), sw1, VolTarg:End Sub
Sub sw12_Hit : vpmTimer.PulseSw(12):PlaySoundAtVol SoundFX("target",DOFTargets), sw1, VolTarg:End Sub
Sub sw51_Hit : vpmTimer.PulseSw(51):PlaySoundAtVol SoundFX("target",DOFTargets), sw1, VolTarg:End Sub
Sub sw52_Hit : vpmTimer.PulseSw(52):PlaySoundAtVol SoundFX("target",DOFTargets), sw1, VolTarg:End Sub

Sub sw16a_Hit : vpmTimer.PulseSw(16) : End Sub

' Drain & Holes
Sub Drain_Hit:playsoundat "drain", Drain:bsTrough.addball me:End Sub

'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep, LRubberstep, RRubberstep

Sub RightSlingShot_Slingshot
    vpmTimer.PulseSw 16
    PlaySoundAt SoundFXDOF("right_slingshot",104,DOFPulse,DOFContactors), sling1
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.TransZ = -30
    RStep = 0
    RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.TransZ = -15
        Case 4:RSLing2.Visible = 0:RSLing3.Visible = 1:sling1.TransZ = 0
        Case 5:RSLing3.Visible = 0:RSLing.Visible = 1:sling1.TransZ = 0:RightSlingShot.TimerEnabled = 0:
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
    vpmTimer.PulseSw 16
    PlaySoundAt SoundFXDOF("left_slingshot",103,DOFPulse,DOFContactors), sling2
    LSling.Visible = 0
    LSling1.Visible = 1
    sling2.TransZ = -30
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.TransZ = -15
        Case 4:LSLing2.Visible = 0:LSLing3.Visible = 1:sling2.TransZ = 0
        Case 5:LSLing3.Visible = 0:LSLing.Visible = 1:sling2.TransZ = 0:LeftSlingShot.TimerEnabled = 0:
    End Select
    LStep = LStep + 1
End Sub

Sub LRubberWall_hit
    LRubber.Visible = 0
    LRubber1.Visible = 1
    LRubberStep = 0
    LRubberWall.TimerEnabled = 1
End Sub

Sub LRubberWall_Timer
    Select Case LRubberStep
        Case 3:LRubber1.Visible = 0:LRubber2.Visible = 1
        Case 4:LRubber2.Visible = 0:LRubber.Visible = 1:LRubberWall.TimerEnabled = 0:
    End Select
    LRubberStep = LRubberStep + 1
End Sub

Sub RRubberWall_hit
    RRubber.Visible = 0
    RRubber1.Visible = 1
    RRubberStep = 0
    RRubberWall.TimerEnabled = 1
End Sub

Sub RRubberWall_Timer
    Select Case RRubberStep
        Case 3:RRubber1.Visible = 0:RRubber2.Visible = 1
        Case 4:RRubber2.Visible = 0:RRubber.Visible = 1:RRubberWall.TimerEnabled = 0:
    End Select
    RRubberStep = RRubberStep + 1
End Sub

Sub sw4_Hit
    PlaySoundAt "metalhit2", sw4
    bsRSaucer.AddBall 0
End Sub

Sub sw54_Hit
    PlaySoundAt "metalhit2", sw54
    bsLSaucer.AddBall 0
End Sub

'******************************
' Diverse Collection Hit Sounds
'******************************

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

Sub Rollovers_Hit (idx)
    PlaySound "fx_sensor", 0, Vol(ActiveBall)*VolRol, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Spinner_Spin
    PlaySoundAtVol "fx_spinner", spinner, VolSpin
End Sub

Sub Rubbers_Hit(idx)
    dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
    If finalspeed > 16 then
        PlaySound "fx_rubber2", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    End if
    If finalspeed >= 6 AND finalspeed <= 16 then
        RandomSoundRubber()
    End If
End Sub

Sub Posts_Hit(idx)
    dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
    If finalspeed > 12 then
        PlaySound "fx_rubber2", 0, Vol(ActiveBall)*VolPo, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    End if
    If finalspeed >= 6 AND finalspeed <= 12 then
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
  NFadeLm 1, GIlamps
    NFadeLm 3, l3
    NFadeLm 3, l3a
    NFadeLm 4, l4
    NFadeLm 4, l4a
    NFadeLm 5, l5
    NFadeLm 5, l5a
    NFadeLm 6, l6
    NFadeLm 6, l6a
    NFadeObjm 6, LeftFlipper_prim, "flipper_topleftON", "flipper_topleft"
    NFadeObjm 6, RightFlipper_prim, "flipper_toprightON", "flipper_topright"
    NFadeLm 7, l7
    NFadeLm 7, l7a
    NFadeLm 8, l8
    NFadeLm 8, l8a
    NFadeLm 12, l12
    NFadeLm 12, l12a
    NFadeLm 13, l13
    NFadeLm 13, l13a
    NFadeLm 14, l14
    NFadeLm 14, l14a
    NFadeLm 15, l15
    NFadeLm 15, l15a
    NFadeLm 16, l16
    NFadeLm 16, l16a
    NFadeLm 17, l17
    NFadeLm 17, l17a
    NFadeLm 18, l18
    NFadeLm 18, l18a
    NFadeLm 19, l19
    NFadeLm 19, l19a
    NFadeLm 20, l20
    NFadeLm 20, l20a
    NFadeLm 21, l21
    NFadeLm 21, l21a
    NFadeLm 22, l22
    NFadeLm 22, l22a
    NFadeLm 23, l23
    NFadeLm 23, l23a
    NFadeLm 24, l24
    NFadeLm 24, l24a
    NFadeLm 25, l25
    NFadeLm 25, l25a
    NFadeLm 26, l26
    NFadeLm 26, l26a
    NFadeLm 27, l27
    NFadeLm 27, l27a
    NFadeLm 28, l28
    NFadeLm 28, l28a
    NFadeLm 29, l29
    NFadeLm 29, l29a
    NFadeLm 30, l30
    NFadeLm 30, l30a
    NFadeLm 31, l31
    NFadeLm 31, l31a
    NFadeLm 32, l32
    NFadeLm 32, l32a
    NFadeLm 33, l33
    NFadeLm 33, l33a
    NFadeLm 34, l34
    NFadeLm 34, l34a
    NFadeLm 35, l35
    NFadeLm 35, l35a
    NFadeLm 36, l36
    NFadeLm 36, l36a
    NFadeLm 37, l37
    NFadeLm 37, l37a
    NFadeLm 38, l38
    NFadeLm 38, l38a
    NFadeLm 39, l39
    NFadeLm 39, l39a
    NFadeLm 40, l40
    NFadeLm 40, l40a
    NFadeLm 41, l41
    NFadeLm 41, l41a
    NFadeLm 43, l43
    NFadeLm 43, l43a
    NFadeLm 44, l44
    NFadeLm 44, l44a
    NFadeLm 45, l45
    NFadeLm 45, l45a
    NFadeLm 46, l46
    NFadeLm 46, l46a
    NFadeLm 47, l47
    NFadeLm 47, l47a
    NFadeLm 51, l51
    NFadeLm 51, l51a
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

 'Reels
Sub FadeReel(nr, reel)
    Select Case FadingLevel(nr)
        Case 2:FadingLevel(nr) = 0
        Case 3:FadingLevel(nr) = 2
        Case 4:reel.Visible = 0:FadingLevel(nr) = 3
        Case 5:reel.Visible = 1:FadingLevel(nr) = 1
    End Select
End Sub

 'Inverted Reels
Sub FadeIReel(nr, reel)
    Select Case FadingLevel(nr)
        Case 2:FadingLevel(nr) = 0
        Case 3:FadingLevel(nr) = 2
        Case 4:reel.Visible = 1:FadingLevel(nr) = 3
        Case 5:reel.Visible = 0:FadingLevel(nr) = 1
    End Select
End Sub

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

Dim LED7(37)
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
LED7(24)=Array()
LED7(25)=Array()
LED7(26)=Array()
LED7(27)=Array()

LED7(28)=Array(d261,d262,d263,d264,d265,d266,d267,LXM,d268)
LED7(29)=Array(d271,d272,d273,d274,d275,d276,d277,LXM,d278)
LED7(30)=Array(d241,d242,d243,d244,d245,d246,d247,LXM,d248)
LED7(31)=Array(d251,d252,d253,d254,d255,d256,d257,LXM,d258)

LED7(32)=Array()
LED7(33)=Array()
LED7(34)=Array()
LED7(35)=Array()
LED7(36)=Array()
LED7(37)=Array()

Sub DisplayTimer_Timer
    Dim ChgLED, II, Num, Chg, Stat, Obj
    ChgLED = Controller.ChangedLEDs(0, &Hffffffff)
    If Not IsEmpty(ChgLED) Then
        For II = 0 To UBound(ChgLED)
            Num = ChgLED(II, 0):Chg = ChgLED(II, 1):Stat = ChgLED(II, 2)
            If Num > 27 Then
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

'Gottlieb Q-Bert's Quest
'added by Inkochnito
Sub editDips
    Dim vpmDips:Set vpmDips = New cvpmDips
    With vpmDips
        .AddForm 700, 400, "Q-Bert's Quest - DIP switches"
        .AddFrame 2, 4, 190, "Maximum credits", 49152, Array("8 credits", 0, "10 credits", 32768, "15 credits", &H00004000, "20 credits", 49152)                                                                                  'dip 15&16
        .AddFrameExtra 2, 80, 190, "Attract Sound", &H000C, Array("off", 0, "every 10 seconds", &H0004, "every 2 minutes", &H0008, "every 4 minutes", &H000C)                                                                     'sounddip 3&4
        .AddFrame 2, 156, 190, "Coin chute 1 and 2 control", &H00002000, Array("seperate", 0, "same", &H00002000)                                                                                                                 'dip 14
        .AddFrame 2, 202, 190, "Playfield special", &H00200000, Array("replay", 0, "extra ball", &H00200000)                                                                                                                      'dip 22
        .AddFrame 2, 248, 190, "Figure 8 shot awards villain lamp", &H00000080, Array("when villain in 1st position only", 0, "when villain in any position", &H00000080)                                                         'dip 8
        .AddFrame 2, 294, 190, "3rd coin chute credits control", &H20000000, Array("no effect", 0, "add 9", &H20000000)                                                                                                           'dip 30
        .AddChkExtra 2, 345, 120, Array("Speech", &H0020)                                                                                                                                                                         'SS-board dip 6
        .AddChkExtra 2, 360, 120, Array("Background sound", &H0010)                                                                                                                                                               'SS-board dip 5
        .AddFrame 205, 4, 190, "High game to date awards", &H00C00000, Array("not displayed and no award", 0, "displayed and no award", &H00800000, "displayed and 2 replays", &H00400000, "displayed and 3 replays", &H00C00000) 'dip 23&24
        .AddFrame 205, 80, 190, "Pyramids needed for special", &HC0000000, Array("5 (3 ball) or 6 (5 ball)", 0, "6            or 7", &H80000000, "7            or 8", &H40000000, "8            or 9", &HC0000000)                'dip 31&32
        .AddFrame 205, 156, 190, "Balls per game", &H01000000, Array("5 balls", 0, "3 balls", &H01000000)                                                                                                                         'dip 25
        .AddFrame 205, 202, 190, "Replay limit", &H04000000, Array("no limit", 0, "one per game", &H04000000)                                                                                                                     'dip 27
        .AddFrame 205, 248, 190, "Novelty", &H08000000, Array("normal", 0, "extra ball and replay scores points", &H08000000)                                                                                                     'dip 28
        .AddFrame 205, 294, 190, "Game mode", &H10000000, Array("replay", 0, "extra ball", &H10000000)                                                                                                                            'dip 29
        .AddChk 205, 345, 120, Array("Match feature", &H02000000)                                                                                                                                                                 'dip 26
        .AddChk 205, 360, 190, Array("Attract mode control (is aways on)", &H00000040)                                                                                                                                            'dip 7
        .AddLabel 50, 380, 300, 20, "After hitting OK, press F3 to reset game with new settings."
    End With
    Dim extra
    extra = Controller.Dip(4) + Controller.Dip(5) * 256
    extra = vpmDips.ViewDipsExtra(extra)
    Controller.Dip(4) = extra And 255
    Controller.Dip(5) = (extra And 65280) \ 256 And 255
End Sub
Set vpmShowDips = GetRef("editDips")

Sub Table1_Exit()
    Controller.Pause = False
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
'    JP's VP10 Rolling Sounds
'*****************************************

Const tnob = 3' total number of balls
ReDim rolling(tnob)
InitRolling

Sub InitRolling
    Dim i
    For i = 0 to tnob
        rolling(i) = FALSE
    Next
End Sub

Sub RollingTimer_Timer()
    Dim BOT, b
    BOT = GetBalls

    ' stop the sound of deleted balls
    For b = UBound(BOT) + 1 to tnob
        rolling(b) = FALSE
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


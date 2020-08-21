' Batman '66 - Williams / 4 Players
' Theme by LuvtheCubs
' VPX by JPSalas May 2016, modded with permission
' LE Sound Mod scripting by LuvtheCubs. Visual LE enhancements by LuvtheCubs
' Based on the script by desktruk & gaston

Option Explicit
Randomize

' Thalamus 2019 March : Improved directional sounds
' !! NOTE : Table not verified yet !!

' Options
' Volume devided by - lower gets higher sound

Const VolDiv = 2000    ' Lower number, louder ballrolling/collition sound
Const VolCol = 10      ' Ball collition divider ( voldiv/volcol )

' The rest of the values are multipliers
'
'  .5 = lower volume
' 1.5 = higher volume

Const VolBump   = 2    ' Bumpers volume.
Const VolGates  = 1    ' Gates volume.
Const VolMetal  = 1    ' Metals volume.
Const VolRB     = 1    ' Rubber bands volume.
Const VolPo     = 1    ' Rubber posts volume.
Const VolPi     = 1    ' Rubber pins volume.
Const VolPlast  = 1    ' Plastics volume.
Const VolTarg   = 1    ' Targets volume.
Const VolWood   = 1    ' Woods volume.
Const VolFlip   = 1    ' Flipper volume.

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

LoadVPM "01120100", "s4.vbs", 3.02

Dim bsTrough, dtRBank, dtLBank, bsLHole, x, plungerIM

Const cGameName = "flash_l1"

Const UseSolenoids = 2
Const UseLamps = 0
Const UseGI = 0
Const UseSync = 0
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
end if

if B2SOn = true then VarHidden = 1

' Standard Sounds
Const SSolenoidOn = "fx_Solenoid"
Const SSolenoidOff = ""
Const SFlipperOn = "fx_FlipperUp"
Const SFlipperOff = "fx_FlipperDown"
Const SCoin = "fx_coin"

'************
' Table init.
'************

Sub Table1_Init
    vpmInit me

    With Controller
        .GameName = cGameName
        .SplashInfoLine = "Flash, The LE (Original 2018) v1.01" & vbNewLine & "VPX table by JPSalas" & vbNewLine & "Mod by Onevox"
        .Games(cGameName).Settings.Value("rol") = 0 '1= rotated display, 0= normal
        .Games(cGameName).Settings.Value("sound") = 0
        .HandleMechanics = 0
        .HandleKeyboard = 0
        .ShowDMDOnly = 1
        .ShowFrame = 0
        .ShowTitle = 0
        .Hidden = VarHidden
        If Err Then MsgBox Err.Description
    End With
    On Error Goto 0
    Controller.SolMask(0) = 0
    vpmTimer.AddTimer 2000, "Controller.SolMask(0)=&Hffffffff'" 'ignore all solenoids - then add the timer to renable all the solenoids after 2 seconds
    Controller.Run

    ' Nudging
    vpmNudge.TiltSwitch = 47
    vpmNudge.Sensitivity = 1
    vpmNudge.TiltObj = Array(Bumper1, Bumper2, Bumper3, leftSlingshot, RightSlingshot)

    ' Trough
    Set bsTrough = New cvpmBallStack
    With bsTrough
        .InitSw 0, 48, 0, 0, 0, 0, 0, 0
        .InitKick BallRelease, 80, 6
        .InitEntrySnd "fx_Solenoid", "fx_Solenoid"
        .InitExitSnd SoundFX("fx_ballrel", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
        .Balls = 1
    End With

    ' Left 5 Drop targets
    set dtLBank = new cvpmdroptarget
    With dtLBank
        .InitDrop Array(sw36, sw35, sw34, sw33, sw32), Array(36, 35, 34, 33, 32)
        .initsnd SoundFX("fx_droptarget", DOFContactors), SoundFX("fx_resetdrop", DOFContactors)
        .AllDownSw = 37
        .CreateEvents "dtLBank"
    End With

    ' Right 3 Drop targets
    set dtRBank = new cvpmdroptarget
    With dtRBank
        .InitDrop Array(sw30, sw29, sw28), Array(30, 29, 28)
        .initsnd SoundFX("fx_droptarget", DOFContactors), SoundFX("fx_resetdrop", DOFContactors)
        .AllDownSw = 31
        .CreateEvents "dtRBank"
    End With

    ' Left Eject Hole
    Set bsLHole = New cvpmBallStack
    With bsLHole
        .InitSaucer sw27, 27, 190, 8
        .InitExitSnd SoundFX("fx_kicker", DOFContactors), SoundFX("fx_kicker", DOFContactors)
        .KickForceVar = 3
        .KickAngleVar = 1
    End With

    ' Main Timer init
    PinMAMETimer.Interval = PinMAMEInterval
    'PinMAMETimer.Enabled = 1:PlaySound"BM02"
	PinMAMETimer.Enabled = 1:PlayMusic
End Sub

Sub Table1_Paused:Controller.Pause = 1:End Sub
Sub Table1_unPaused:Controller.Pause = 0:End Sub

'**********
' Keys
'**********

Sub table1_KeyDown(ByVal Keycode)
    If keycode = LeftTiltKey Then Nudge 90, 5:PlaySound SoundFX("fx_nudge", 0), 0, 1, -0.1, 0.25
    If keycode = RightTiltKey Then Nudge 270, 5:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0.1, 0.25
    If keycode = CenterTiltKey Then Nudge 0, 6:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0, 0.25
    If keycode = PlungerKey Then PlaySoundAtVol "fx_PlungerPull", Plunger, 1 :Plunger.Pullback
    If vpmKeyDown(keycode) Then Exit Sub
    If keycode = KeyRules Then Rules
End Sub

Sub table1_KeyUp(ByVal Keycode)
    If keycode = 2 Then StopSound"BM02"
    If keycode = PlungerKey Then PlaySoundAtVol "fx_plunger", Plunger, 1:Plunger.Fire
    If vpmKeyUp(keycode) Then Exit Sub
End Sub

'*********
' Switches
'*********

' Slings & Rubbers
' Slings
Dim LStep, RStep
Dim AA
AA=0

Sub LeftSlingShot_Slingshot
    PlaySoundAtVol SoundFX("fx_slingshot", DOFContactors), Lemk, 1
    PlaySound"FL16"
    LeftSling4.Visible = 1
    Lemk.RotX = 26
    LStep = 0
    vpmTimer.PulseSw 41
    LeftSlingShot.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 1:LeftSLing4.Visible = 0:LeftSLing3.Visible = 1:Lemk.RotX = 14
        Case 2:LeftSLing3.Visible = 0:LeftSLing2.Visible = 1:Lemk.RotX = 2
        Case 3:LeftSLing2.Visible = 0:Lemk.RotX = -10:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

Sub RightSlingShot_Slingshot
    PlaySoundAtVol SoundFX("fx_slingshot", DOFContactors), Remk, 1
    PlaySound"FL16"
    RightSling4.Visible = 1
    Remk.RotX = 26
    RStep = 0
    vpmTimer.PulseSw 42
    RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 1:RightSLing4.Visible = 0:RightSLing3.Visible = 1:Remk.RotX = 14
        Case 2:RightSLing3.Visible = 0:RightSLing2.Visible = 1:Remk.RotX = 2
        Case 3:RightSLing2.Visible = 0:Remk.RotX = -10:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

Sub sw24_Hit:vpmTimer.PulseSw 24:PlaySound"FL04":PlaySound "fx_rubber", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0 , AudioFade(ActiveBall):End Sub
Sub sw25_Hit:vpmTimer.PulseSw 25:PlaySound"FL04":PlaySound "fx_rubber", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0 , AudioFade(ActiveBall):End Sub
Sub sw26_Hit:vpmTimer.PulseSw 26:PlaySound"FL04":PlaySound "fx_rubber", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0 , AudioFade(ActiveBall):End Sub
Sub sw16_Hit:vpmTimer.PulseSw 16:PlaySound"FL04":PlaySound "fx_rubber", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0 , AudioFade(ActiveBall):End Sub
Sub sw11_Hit:vpmTimer.PulseSw 11:PlaySound"FL04":PlaySound "fx_rubber", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0 , AudioFade(ActiveBall):End Sub
Sub sw38_Hit:vpmTimer.PulseSw 38:PlaySound"FL04":PlaySound "fx_rubber", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0 , AudioFade(ActiveBall):End Sub
Sub sw38a_Hit:vpmTimer.PulseSw 38:PlaySound"FL04":PlaySound "fx_rubber", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub

' Bumpers
Sub Bumper1_Hit:vpmTimer.PulseSw 18:PlaySound "FL14":PlaySoundAtVol SoundFX("fx_bumper", DOFContactors), ActiveBall, VolBump:End Sub
Sub Bumper2_Hit:vpmTimer.PulseSw 19:PlaySound "FL14":PlaySoundAtVol SoundFX("fx_bumper", DOFContactors), ActiveBall, VolBump:End Sub
Sub Bumper3_Hit:vpmTimer.PulseSw 20:PlaySound "FL14":PlaySoundAtVol SoundFX("fx_bumper", DOFContactors), ActiveBall, VolBump:End Sub

' Spinners

Sub sw17_Spin:vpmTimer.PulseSw 17:PlaySound"FL13":PlaySoundAtVol "spinner", sw17, 1 :End Sub

' Drain & holes
Sub Drain_Hit:AA=0:BIP.enabled=False:StopSound"BM00":PlaySound"FL01":PlaysoundAtVol "fx_drain", Drain, 1:bsTrough.AddBall Me:End Sub
Sub sw27_Hit:PlaySound"BM07":PlaySoundAtVol "fx_kicker_enter", sw27, 1 :PlaySound "BM07":sw27.timerinterval=2500:sw27.timerenabled=true:bsLHole.AddBall Me:End Sub


' Rollovers
Sub sw43_Hit:Controller.Switch(43) = 1:PlaySound"FL03":PlaySound "fx_sensor", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub sw43_UnHit:Controller.Switch(43) = 0:End Sub

Sub sw44_Hit:Controller.Switch(44) = 1:PlaySound"FL03":PlaySound "fx_sensor", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), AudioFade(ActiveBall):End Sub
Sub sw44_UnHit:Controller.Switch(44) = 0:End Sub

  Sub sw45_Hit:Controller.Switch(45) = 1:PlaySound"FL03":PlaySound "fx_sensor", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), AudioFade(ActiveBall):End Sub
Sub sw45_UnHit:Controller.Switch(45) = 0:End Sub

Sub sw46_Hit:Controller.Switch(46) = 1:PlaySound"FL03":PlaySound "fx_sensor", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), AudioFade(ActiveBall):End Sub
Sub sw46_UnHit:Controller.Switch(46) = 0:End Sub

Sub sw12_Hit:Controller.Switch(12) = 1:PlaySound"FL10":PlaySound "fx_sensor", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), AudioFade(ActiveBall):End Sub
Sub sw12_UnHit:Controller.Switch(12) = 0:End Sub

Sub sw13_Hit:Controller.Switch(13) = 1:PlaySound"FL10":PlaySound "fx_sensor", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), AudioFade(ActiveBall):End Sub
Sub sw13_UnHit:Controller.Switch(13) = 0:End Sub

Sub sw14_Hit:Controller.Switch(14) = 1:PlaySound"FL10":PlaySound "fx_sensor", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), AudioFade(ActiveBall):End Sub
Sub sw14_UnHit:Controller.Switch(14) = 0:End Sub

Sub sw15_Hit:Controller.Switch(15) = 1:PlaySound"FL10":PlaySound "fx_sensor", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), AudioFade(ActiveBall):End Sub
Sub sw15_UnHit:Controller.Switch(15) = 0:End Sub

Sub sw9_Hit:Controller.Switch(9) = 1:Playsound"FL07":PlaySound "fx_sensor", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), AudioFade(ActiveBall):End Sub
Sub sw9_UnHit:Controller.Switch(9) = 0:End Sub

Sub sw21_Hit:Controller.Switch(21) = 1:Playsound"FL07":PlaySound "fx_sensor", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), AudioFade(ActiveBall):End Sub
Sub sw21_UnHit:Controller.Switch(21) = 0:End Sub

Sub sw22_Hit:Controller.Switch(22) = 1:Playsound"FL07":PlaySound "fx_sensor", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), AudioFade(ActiveBall):End Sub
Sub sw22_UnHit:Controller.Switch(22) = 0:End Sub

Sub sw23_Hit:Controller.Switch(23) = 1:Playsound"FL07":PlaySound "fx_sensor", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), AudioFade(ActiveBall):End Sub
Sub sw23_UnHit:Controller.Switch(23) = 0:End Sub

Sub sw10_Hit:Controller.Switch(10) = 1:Playsound"FL07":PlaySound "fx_sensor", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), AudioFade(ActiveBall):End Sub
Sub sw10_UnHit:Controller.Switch(10) = 0:End Sub

' Targets

Sub sw28_Hit:If l30.state = 1 then Playsound"BM04" Else Playsound"BM04":End If:End Sub

Sub sw29_Hit:If l31.state = 1 then Playsound"BM05" Else Playsound"BM05":End If:End Sub

Sub sw30_Hit:If l32.state = 1 then Playsound"BM06" Else Playsound"BM06":End If:End Sub


Sub sw32_Hit:If l29.state = 1 then Playsound"BM04" Else Playsound"BM04":End If:End Sub
Sub sw33_Hit:If l28.state = 1 then Playsound"BM05" Else Playsound"BM05":End If:End Sub
Sub sw34_Hit:If l27.state = 1 then Playsound"BM06" Else Playsound"BM06":End If:End Sub
Sub sw35_Hit:If l26.state = 1 then Playsound"BM04" Else Playsound"BM04":End If:End Sub
Sub sw36_Hit:If l25.state = 1 then Playsound"BM05" Else Playsound"BM05":End If:End Sub

Sub sw39_Hit:vpmTimer.PulseSw 39:PlaySoundAtVol SoundFX("fx_target", DOFContactors), ActiveBall, VolTarg:End Sub
Sub sw40_Hit:vpmTimer.PulseSw 40:PlaySoundAtVol SoundFX("fx_target", DOFContactors), ActiveBall, VolTarg:End Sub


'Sub Gate1_Hit():AA=1:StopSound"BM08":StopSound"BM09":StopSound"BM10":StopSound"BM11":StopSound"BM12":PlaySound "BM00":BIP.enabled=True:BIP.interval=115300:End Sub
Sub Gate1_Hit():AA=1:StopMusic():PlaySound "BM00":BIP.enabled=True:BIP.interval=115300:End Sub

Sub BIP_Timer
BIP.enabled=False
PlaySound"BM00"
BIP.enabled=True:BIP.interval=115300
End Sub

' Random Music

Sub PlayMusic()
Dim RndSound
RndSound = Rnd(1)*4
Select Case (Int(RndSound))
Case 0: PlaySound"BM08"
Case 1: PlaySound"BM09"
Case 2: PlaySound"BM10"
Case 3: PlaySound"BM11"
Case 4: PlaySound"BM12"
End Select
End Sub

Sub StopMusic()
StopSound"BM08"
StopSound"BM09"
StopSound"BM10"
StopSound"BM11"
StopSound"BM12"
End Sub

'*********
'Solenoids
'*********
SolCallback(1) = "bsTrough.SolOut"
SolCallback(2) = "SolReset 1,"
SolCallback(3) = "SolReset 2,"
SolCallback(4) = "dtRBank.SolDropUp"
SolCallback(5) = "bsLHole.SolOut"
solcallback(14) = "vpmsolsound SoundFX(""fx_knocker"",DOFKnocker),"
solcallback(23) = "SolRun"

'**************
' Flipper Subs
'**************

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
    If Enabled Then
        PlaySoundAtVol SoundFX("fx_flipperup", DOFContactors), LeftFlipper, VolFlip
        LeftFlipper.RotateToEnd
    Else
        PlaySoundAtVol SoundFX("fx_flipperdown", DOFContactors), LeftFlipper, VolFlip
        LeftFlipper.RotateToStart
    End If
End Sub

Sub SolRFlipper(Enabled)
    If Enabled Then
        PlaySoundAtVol SoundFX("fx_flipperup", DOFContactors), RightFlipper, VolFlip
        RightFlipper.RotateToEnd
        RightFlipper2.RotateToEnd
    Else
        PlaySoundAtVol SoundFX("fx_flipperdown", DOFContactors), RightFlipper, VolFlip
        RightFlipper.RotateToStart
        RightFlipper2.RotateToStart
    End If
End Sub

Sub LeftFlipper_Collide(parm)
    PlaySoundAtBallVol "fx_rubber_flipper", parm / 10
End Sub

Sub Rightflipper2_Collide(parm)
    PlaySoundAtBallVol "fx_rubber_flipper", parm / 10
End Sub

Sub Rightflipper_Collide(parm)
    PlaySoundAtBallVol "fx_rubber_flipper", parm / 10
End Sub

'Solenoid subs

Sub SolReset(No, Enabled)
    If Enabled Then
        Controller.Switch(37) = False
        If no < 2 Then
            dtLbank.SolUnHit 1, True
            dtLbank.SolUnHit 2, True
            dtLbank.SolUnHit 3, True
        Else
            dtLbank.SolUnHit 4, True
            dtLbank.SolUnHit 5, True
        End If
    End If
End Sub

Sub SolRun(Enabled)
    vpmNudge.SolGameOn Enabled
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
LampTimer.Interval = 10 'lamp fading speed
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
	UpdateLeds
    UpdateLamps
End Sub

 Sub UpdateLamps
    NFadeL 2, l2
    NFadeL 3, l3
    NFadeL 4, l4
    NFadeL 5, l5
    NFadeL 6, l6
    NFadeL 7, l7
    NFadeL 8, l8
    NFadeL 9, l9
    NFadeL 10, l10
    NFadeL 12, l12
    NFadeL 13, l13
    NFadeL 14, l14
    NFadeL 15, l15
    NFadeL 16, l16
    NFadeL 17, l17
    NFadeL 18, l18
    NFadeL 19, l19
    NFadeL 20, l20
    NFadeL 21, l21
    NFadeL 22, l22
    NFadeL 23, l23
    NFadeL 24, l24
    NFadeL 25, l25
    NFadeL 26, l26
    NFadeL 27, l27
    NFadeL 28, l28
    NFadeL 29, l29
    NFadeL 30, l30
    NFadeL 31, l31
    NFadeL 32, l32
    NFadeL 33, l33
    NFadeL 34, l34
    NFadeL 35, l35
    NFadeL 36, l36
    NFadeL 37, l37
    NFadeL 38, l38
    NFadeL 39, l39
    NFadeLm 43, l43b
    NFadeLm 43, l43e
    NFadeLm 43, l43g
    NFadeL 43, l43
    NFadeLm 44, l44b
    NFadeLm 44, l44c
    NFadeL 44, l44
    NFadeLm 45, l45b
    NFadeL 45, l45


	NFadeLm 40, bumper3l
	NFadeL 40, bumper3l1
	NFadeLm 41, bumper1l
	NFadeL 41, bumper1l1
	NFadeLm 42, bumper2l
	NFadeL 42, bumper2l1

   'Backdrop lights
	NFadeL 50, l50     '50 1 can play
	NFadeL 51, l51     '51 2 can play
	NFadeL 52, l52     '53 3 can play
	NFadeL 53, l53     '54 4 can play
	NFadeL 57, l57     '57 1 player
	NFadeL 58, l58     '58 2 player
	NFadeL 59, l59     '59 3 player
	NFadeL 60, l60     '60 4 player

	NFadeT 54, l54, "MATCH"
	NFadeT 55, l55, "BALL" & vbNewLine & "In PLAY"
	NFadeT 61, l61, "TILT"
	NFadeT 62, l62, "GAME OVER"
	NFadeT 63, l63, "SAME PLAYER" & vbNewLine & "SHOOTS AGAIN"
	NFadeT 64, l64, "HIGHEST SCORE"
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
		FlashRepeat(x) = 20		' how many times the flash repeats
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

Sub Flashm(nr, object) 'multiple flashers, it doesn't change anything, it just follows the main flasher
    Select Case FadingLevel(nr)
        Case 4,5
    Object.IntensityScale = FlashLevel(nr)
End Select
End Sub

Sub FlashBlink(nr, object)
    Select Case FadingLevel(nr)
        Case 4 'off
            FlashLevel(nr) = FlashLevel(nr) - FlashSpeedDown(nr)
            If FlashLevel(nr) < FlashMin(nr) Then
                FlashLevel(nr) = FlashMin(nr)
                FadingLevel(nr) = 0 'completely off
            End if
            Object.IntensityScale = FlashLevel(nr)
            If FadingLevel(nr) = 0 AND FlashRepeat(nr) Then                          'repeat the flash
                FlashRepeat(nr) = FlashRepeat(nr) -1
				If FlashRepeat(nr) Then FadingLevel(nr) = 5
            End If
        Case 5 ' on
            FlashLevel(nr) = FlashLevel(nr) + FlashSpeedUp(nr)
            If FlashLevel(nr) > FlashMax(nr) Then
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
        Case 6, 7, 8:FadingLevel(nr) = FadingLevel(nr) + 1              'wait
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

Sub NFadeTm(nr, object, b)
    Select Case FadingLevel(nr)
        Case 4:object.Text = ""
        Case 5:object.Text = message
    End Select
End Sub

'************************************
'          LEDs Display
'************************************

Dim Digits(28)

Set Digits(0) = a1
Set Digits(1) = a2
Set Digits(2) = a3
Set Digits(3) = a4
Set Digits(4) = a5
Set Digits(5) = a6

Set Digits(6) = b1
Set Digits(7) = b2
Set Digits(8) = b3
Set Digits(9) = b4
Set Digits(10) = b5
Set Digits(11) = b6

Set Digits(12) = c1
Set Digits(13) = c2
Set Digits(14) = c3
Set Digits(15) = c4
Set Digits(16) = c5
Set Digits(17) = c6

Set Digits(18) = d1
Set Digits(19) = d2
Set Digits(20) = d3
Set Digits(21) = d4
Set Digits(22) = d5
Set Digits(23) = d6

Set Digits(24) = e1
Set Digits(25) = e2
Set Digits(26) = e3
Set Digits(27) = e4

Sub UPdateLEDs
	On Error Resume Next
	Dim ChgLED, ii, jj, chg, stat
	ChgLED = Controller.ChangedLEDs(&H0000003f, &Hffffffff)
	If Not IsEmpty(ChgLED) Then
		For ii = 0 To UBound(ChgLED)
			chg = chgLED(ii, 1):stat = chgLED(ii, 2)

			Select Case stat
				Case 0:Digits(chgLED(ii, 0) ).SetValue 0    'empty
				Case 63:Digits(chgLED(ii, 0) ).SetValue 1   '0
				Case 6:Digits(chgLED(ii, 0) ).SetValue 2    '1
				Case 91:Digits(chgLED(ii, 0) ).SetValue 3   '2
				Case 79:Digits(chgLED(ii, 0) ).SetValue 4   '3
				Case 102:Digits(chgLED(ii, 0) ).SetValue 5  '4
				Case 109:Digits(chgLED(ii, 0) ).SetValue 6  '5
				Case 124:Digits(chgLED(ii, 0) ).SetValue 7  '6
				Case 125:Digits(chgLED(ii, 0) ).SetValue 7  '6
				Case 252:Digits(chgLED(ii, 0) ).SetValue 7  '6
				Case 7:Digits(chgLED(ii, 0) ).SetValue 8    '7
				Case 127:Digits(chgLED(ii, 0) ).SetValue 9  '8
				Case 103:Digits(chgLED(ii, 0) ).SetValue 10 '9
				Case 111:Digits(chgLED(ii, 0) ).SetValue 10 '9
				Case 231:Digits(chgLED(ii, 0) ).SetValue 10 '9
				Case 128:Digits(chgLED(ii, 0) ).SetValue 0  'empty
				Case 191:Digits(chgLED(ii, 0) ).SetValue 1  '0
				Case 832:Digits(chgLED(ii, 0) ).SetValue 2  '1
				Case 896:Digits(chgLED(ii, 0) ).SetValue 2  '1
				Case 768:Digits(chgLED(ii, 0) ).SetValue 2  '1
				Case 134:Digits(chgLED(ii, 0) ).SetValue 2  '1
				Case 219:Digits(chgLED(ii, 0) ).SetValue 3  '2
				Case 207:Digits(chgLED(ii, 0) ).SetValue 4  '3
				Case 230:Digits(chgLED(ii, 0) ).SetValue 5  '4
				Case 237:Digits(chgLED(ii, 0) ).SetValue 6  '5
				Case 253:Digits(chgLED(ii, 0) ).SetValue 7  '6
				Case 135:Digits(chgLED(ii, 0) ).SetValue 8  '7
				Case 255:Digits(chgLED(ii, 0) ).SetValue 9  '8
				Case 239:Digits(chgLED(ii, 0) ).SetValue 10 '9
			End Select
		Next
	End IF
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

Const tnob = 1 ' total number of balls in this table is 4, but always use a higher number here because of the timing
ReDim rolling(tnob)
InitRolling

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

'******************
' RealTime Updates
'******************

Set MotorCallback = GetRef("RealTimeUpdates")

Sub RealTimeUpdates
    RollingUpdate
    GIUpdate
End Sub

Sub GiON
    For each x in aGiLights
        x.State = 1
    Next
End Sub

Sub GiOFF
    For each x in aGiLights
        x.State = 0
    Next
End Sub

Dim OldGiState
OldGiState = -1 'start with the Gi off

Sub GIUpdate
    Dim tmp, obj
    tmp = Getballs
    If UBound(tmp) <> OldGiState Then
        OldGiState = Ubound(tmp)
        If UBound(tmp) = -1 Then
            GiOff
			PlayMusic
        Else
            GiOn
        End If
    End If
End Sub

'******************************
' Diverse Collection Hit Sounds
'******************************

Sub aMetals_Hit(idx):PlaySound "fx_metalhit", 0, Vol(ActiveBall)*VolMetal, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aRubber_Bands_Hit(idx):PlaySound "fx_rubber_band", 0, Vol(ActiveBall)*VolRB, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aRubber_Posts_Hit(idx):PlaySound "fx_rubber", 0, Vol(ActiveBall)*VolPo, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aRubber_Pins_Hit(idx):PlaySound "fx_postrubber", 0, Vol(ActiveBall)*VolPi, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aPlastics_Hit(idx):PlaySound "fx_PlasticHit", 0, Vol(ActiveBall)*VolPlast, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aGates_Hit(idx):PlaySound "fx_Gate", 0, Vol(ActiveBall)*VolGates, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aWoods_Hit(idx):PlaySound "fx_Woodhit", 0, Vol(ActiveBall)*VolWood, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub

'******
' Rules
'******
Dim Msg(17)
Sub Rules()
    Msg(0) = "Batman - by LuvtheCubs" &Chr(10) &Chr(10)
    Msg(1) = ""
    Msg(2) = "Insert coin and wait for machine to reset before inserting coin for"
    Msg(3) = " next player."
    Msg(4) = ""
    Msg(5) = "Making 1 - 2 - 3  lights 2x. Making 1 - 2 - 3 - 4  lights 3x."
    Msg(6) = "Making 3 bank drop targets advances thru 5, 10, and"
    Msg(7) = "20 Thousand."
    Msg(8) = ""
    Msg(9) = "Making 5 bank targets 1st time advances Hole Kicker value, 2nd"
    Msg(10) = "time lights Extra Ball, 3rd time lights out lane specials."
    Msg(11) = ""
    Msg(12) = "Tilt penalty - Ball in Play - does not disqualify player."
    Msg(13) = "Special scores 1 Credit."
    Msg(14) = "Beating highest score scores 3 credits."
    Msg(15) = "Matching last two numbers on score with numbers in match window"
    Msg(16) = "on backglass scores 1 credit"
	Msg(17) = ""
    Msg(18) = "Difusing bomb 3 times scores 50000"

    For X = 1 To 18
        Msg(0) = Msg(0) + Msg(X) &Chr(13)
    Next
    MsgBox Msg(0), , "         Instructions and Rule Card"
End Sub

Sub Table1_Exit():Controller.Games(cGameName).Settings.Value("sound") = 1:Controller.Stop:End Sub

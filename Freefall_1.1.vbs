' Freefall / IPD No. 953 / January, 1981 / 4 Players
' http://www.ipdb.org/machine.cgi?id=953
' VPX v1.0 by JPSalas 2017
' Many parts of the script inspired/copied from the old table by TAB/Destruk/LuvThatApex/Inkochnito

' Thalamus 2018-07-23
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Changed UseSolenoids=1 to 2
' No special SSF tweaks yet.
' This is a JP table. He often uses walls as switches so I need to be careful of using PlaySoundAt

Option Explicit
Randomize

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

LoadVPM "01200000", "stern.vbs", 3.02

Dim bsTrough, dtDrop5, dtDrop3, plSkyway, plSkyraider
Dim x

Const cGameName = "freefall" ' freefall rom
'Const cGameName = "freefafp" ' freefall freeplay

Const UseSolenoids = 2
Const UseLamps = 1
Const UseGI = 0
Const UseSync = 0 'set it to 1 if the table runs too fast
Const HandleMech = 0

' Standard Sounds
Const SSolenoidOn = "fx_Solenoidon"
Const SSolenoidOff = "fx_SolenoidOff"
Const SCoin = "fx_coin"

Sub table1_Init
    vpmInit me
    vpmMapLights aLights ' Map all lamps into lights array

    With Controller
        .GameName = cGameName
        .SplashInfoLine = "Freefall - Stern 1991" & vbNewLine & "VPX-VPM table by JPSalas v.1.0"
        .Games(cGameName).Settings.Value("rol") = 0
        .HandleMechanics = 0
        .HandleKeyboard = 0
        .ShowDMDOnly = 1
        .ShowFrame = 0
        .ShowTitle = 0
        .Hidden = 0
        If Err Then MsgBox Err.Description
    End With
    On Error Goto 0
    Controller.SolMask(0) = 0
    vpmTimer.AddTimer 2000, "Controller.SolMask(0)=&Hffffffff'" 'ignore all solenoids - then add the Timer to renable all the solenoids after 2 seconds
    Controller.Run GetPlayerHWnd

    ' Nudging
    vpmNudge.TiltSwitch = 7
    vpmNudge.Sensitivity = 5
    vpmNudge.TiltObj = Array(Bumper1, Bumper2, Bumper3, LeftSlingshot, RightSlingshot)

    ' Trough & Ball Release
    Set bsTrough = New cvpmBallStack
    With bsTrough
        .InitSw 8, 33, 34, 35, 0, 0, 0, 0
        .InitKick BallRelease, 180, 10
        .InitExitSnd SoundFX("fx_ballrel", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
        .InitEntrySnd "fx_solenoid", "fx_solenoid"
        .IsTrough = True
        .Balls = 3
    End With

    ' 5 Droptargets Left
    Set dtDrop5 = New cvpmDropTarget
    With dtDrop5
        .InitDrop Array(sw21, sw20, sw19, sw18, sw17), Array(21, 20, 19, 18, 17)
        .initsnd SoundFX("fx_droptarget", DOFDropTargets), SoundFX("fx_resetdrop", DOFContactors)
    '.CreateEvents "dtDrop5"  ' we do it manually, because of the new droptargets
    End With

    ' 3 Droptargets Center
    Set dtDrop3 = New cvpmDropTarget
    With dtDrop3
        .InitDrop Array(sw24, sw23, sw22), Array(24, 23, 22)
        .initsnd SoundFX("fx_droptarget", DOFDropTargets), SoundFX("fx_resetdrop", DOFContactors)
    '.CreateEvents "dtDrop3"  ' we do it manually, because of the new droptargets
    End With

    ' Skyway Impulse Plunger
    Const IMPowerSetting = 42 'Plunger Power
    Const IMTime = 0.6        ' Time in seconds for Full Plunge
    Set plSkyway = New cvpmImpulseP
    With plSkyway
        .InitImpulseP sw5, IMPowerSetting, IMTime
        .Random 0.3
        .switch 5
        .InitExitSnd SoundFX("fx_popper", DOFContactors), SoundFX("fx_popper", DOFContactors)
        .CreateEvents "plSkyway"
    End With

    ' Skyraider Impulse Plunger
    Const IMPowerSetting2 = 42 'Plunger Power
    Const IMTime2 = 0.6        ' Time in seconds for Full Plunge
    Set plSkyraider = New cvpmImpulseP
    With plSkyraider
        .InitImpulseP sw4, IMPowerSetting2, IMTime2
        .Random 0.3
        .switch 4
        .InitExitSnd SoundFX("fx_popper", DOFContactors), SoundFX("fx_popper", DOFContactors)
        .CreateEvents "plSkyraider"
    End With

    ' Main Timer init
    PinMAMETimer.Interval = PinMAMEInterval
    PinMAMETimer.Enabled = 1

	' Remove the cabinet rails if in FS mode
	If Table1.ShowDT = False then
		lrail.Visible = False
		rrail.Visible = False
	End If
End Sub

'**********
' Keys
'**********

Sub table1_KeyDown(ByVal Keycode)
    If keycode = LeftTiltKey Then Nudge 90, 5:PlaySound SoundFX("fx_nudge", 0), 0, 1, -0.1, 0.25
    If keycode = RightTiltKey Then Nudge 270, 5:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0.1, 0.25
    If keycode = CenterTiltKey Then Nudge 0, 6:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0, 0.25
    If keycode = PlungerKey Then PlaySound "fx_PlungerPull", 0, 1, 0.1, 0.05:Plunger.Pullback
    If vpmKeyDown(keycode) Then Exit Sub
End Sub

Sub table1_KeyUp(ByVal Keycode)
    If vpmKeyUp(keycode) Then Exit Sub
    If keycode = PlungerKey Then PlaySound "fx_plunger", 0, 1, 0.1, 0.05:Plunger.Fire
End Sub

'*********
' Switches
'*********

' Slings
Dim LStep, RStep

Sub LeftSlingShot_Slingshot
    PlaySound SoundFX("fx_slingshot", DOFContactors), 0, 1, -0.05, 0.05
    LeftSling4.Visible = 1
    Lemk.RotX = 26
    LStep = 0
    vpmTimer.PulseSw 13
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
    PlaySound SoundFX("fx_slingshot", DOFContactors), 0, 1, 0.05, 0.05
    RightSling4.Visible = 1
    Remk.RotX = 26
    RStep = 0
    vpmTimer.PulseSw 12
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

Sub sw11_Hit:PlaySound "fx_rubber":vpmTimer.PulseSw 11:End Sub

' Bumpers
Sub Bumper1_Hit:vpmTimer.PulseSw 16:PlaySound SoundFX("fx_bumper", DOFContactors), 0, 1, 0, 0.15:End Sub
Sub Bumper2_Hit:vpmTimer.PulseSw 15:PlaySound SoundFX("fx_bumper", DOFContactors), 0, 1, -0.1, 0.15:End Sub
Sub Bumper3_Hit:vpmTimer.PulseSw 14:PlaySound SoundFX("fx_bumper", DOFContactors), 0, 1, 0.1, 0.15:End Sub

'droptargets
Sub sw21_Dropped():dtDrop5.Hit 1:End Sub
Sub sw20_Dropped():dtDrop5.Hit 2:End Sub
Sub sw19_Dropped():dtDrop5.Hit 3:End Sub
Sub sw18_Dropped():dtDrop5.Hit 4:End Sub
Sub sw17_Dropped():dtDrop5.Hit 5:End Sub
Sub sw24_Dropped():dtDrop3.Hit 1:End Sub
Sub sw23_Dropped():dtDrop3.Hit 2:End Sub
Sub sw22_Dropped():dtDrop3.Hit 3:End Sub

'only hit sound
Sub sw21_Hit():PlaySound SoundFX("fx_droptarget", DOFDropTargets), 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
Sub sw20_Hit():PlaySound SoundFX("fx_droptarget", DOFDropTargets), 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
Sub sw19_Hit():PlaySound SoundFX("fx_droptarget", DOFDropTargets), 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
Sub sw18_Hit():PlaySound SoundFX("fx_droptarget", DOFDropTargets), 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
Sub sw17_Hit():PlaySound SoundFX("fx_droptarget", DOFDropTargets), 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
Sub sw24_Hit():PlaySound SoundFX("fx_droptarget", DOFDropTargets), 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
Sub sw23_Hit():PlaySound SoundFX("fx_droptarget", DOFDropTargets), 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
Sub sw22_Hit():PlaySound SoundFX("fx_droptarget", DOFDropTargets), 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub

' Spinners

Sub Spinner1_Spin:vpmTimer.PulseSw 36:PlaySound "fx_spinner", 0, 1, -0.1, 0.15:End Sub

' Eject holes
Sub Drain_Hit:Playsound "fx_drain":bsTrough.AddBall Me:End Sub
Sub vuk_hit:vuk.kick 0, 20, 1.56:End Sub

' Rollovers
Sub sw37_Hit:DOF 113, DOFOn:Controller.Switch(37) = 1:PlaySound "fx_sensor", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
Sub sw37_UnHit:DOF 113, DOFOff:Controller.Switch(37) = 0:End Sub

Sub sw9_Hit:DOF 117, DOFOn:Controller.Switch(9) = 1:PlaySound "fx_sensor", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
Sub sw9_UnHit:DOF 117, DOFOff:Controller.Switch(9) = 0:End Sub

Sub sw40_Hit:DOF 115, DOFOn:Controller.Switch(40) = 1:PlaySound "fx_sensor", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
Sub sw40_UnHit:DOF 115, DOFOff:Controller.Switch(40) = 0:End Sub

Sub sw40a_Hit:DOF 116, DOFOn:Controller.Switch(40) = 1:PlaySound "fx_sensor", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
Sub sw40a_UnHit:DOF 116, DOFOff:Controller.Switch(40) = 0:End Sub

Sub sw9a_Hit:DOF 118, DOFOn:Controller.Switch(9) = 1:PlaySound "fx_sensor", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
Sub sw9a_UnHit:DOF 118, DOFOff:Controller.Switch(9) = 0:End Sub

Sub sw37a_Hit:DOF 114, DOFOn:Controller.Switch(37) = 1:PlaySound "fx_sensor", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
Sub sw37a_UnHit:DOF 114, DOFOff:Controller.Switch(37) = 0:End Sub

Sub sw29_Hit:DOF 110, DOFOn:Controller.Switch(29) = 1:PlaySound "fx_sensor", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
Sub sw29_UnHit:DOF 110, DOFOff:Controller.Switch(29) = 0:End Sub

Sub sw28_Hit:Controller.Switch(28) = 1:PlaySound "fx_sensor", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
Sub sw28_UnHit:Controller.Switch(28) = 0:End Sub

Sub sw27_Hit:Controller.Switch(27) = 1:PlaySound "fx_sensor", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
Sub sw27_UnHit:Controller.Switch(27) = 0:End Sub

Sub sw26_Hit:DOF 112, DOFOn:Controller.Switch(26) = 1:PlaySound "fx_sensor", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
Sub sw26_UnHit:DOF 112, DOFOff:Controller.Switch(26) = 0:End Sub

' Targets

Sub sw30_Hit:vpmTimer.PulseSw 30:PlaySound SoundFX("fx_target", DOFTargets), 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
Sub sw31_Hit:vpmTimer.PulseSw 31:PlaySound SoundFX("fx_target", DOFTargets), 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
Sub sw32_Hit:vpmTimer.PulseSw 32:PlaySound SoundFX("fx_target", DOFTargets), 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
Sub sw35_Hit:vpmTimer.PulseSw 35:PlaySound SoundFX("fx_target", DOFTargets), 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
Sub sw29b_Hit:vpmTimer.PulseSw 29:PlaySound SoundFXDOF("fx_target",111,DOFPulse,DOFTargets), 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
Sub sw31b_Hit:vpmTimer.PulseSw 31:PlaySound SoundFXDOF("fx_target",111,DOFPulse,DOFTargets), 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
Sub sw25_Hit:vpmTimer.PulseSw 25:PlaySound SoundFX("fx_target", DOFTargets), 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
Sub sw26b_Hit:vpmTimer.PulseSw 26:PlaySound SoundFXDOF("fx_target",111,DOFPulse,DOFTargets), 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub

'********************
' 	  Solenoids
'********************

'SolCallback(1)="vpmSolSound ""bumper"","
'SolCallback(2)="vpmSolSound ""bumper"","
'SolCallback(3)="vpmSolSound ""bumper"","
'SolCallback(4)="vpmSolSound ""sling"","
'SolCallback(12)="vpmSolSound ""sling"","

SolCallback(6) = "vpmsolsound SoundFX(""fx_knocker"",DOFKnocker),"
SolCallback(7) = "SolSkyway"
SolCallback(8) = "SolBallRelease"
SolCallback(9) = "SolSkyRaider"
SolCallback(10) = "dtDrop5.SolDropUp"
SolCallback(11) = "dtDrop3.SolDropUp"
SolCallback(13) = "SolOutHole"
SolCallback(17) = "SolBallLock"
SolCallback(19) = "vpmNudge.SolGameOn"
SolCallback(20) = "SolDiv"

Dim sw5Step, sw4Step

Sub SolSkyway(Enabled)
    If Enabled Then
        plSkyway.AutoFire
        sw5Step = 0
        Remk2.RotX = 26
        sw5t.TimerEnabled = 1
    End If
End Sub

Sub sw5t_Timer
    Select Case sw5Step
        Case 1:Remk2.RotX = 14
        Case 2:Remk2.RotX = 2
        Case 3:Remk2.RotX = -10:sw5t.TimerEnabled = 0
    End Select

    sw5Step = sw5Step + 1
End Sub

Sub SolSkyraider(Enabled)
    If Enabled Then
        plSkyraider.AutoFire
        sw4Step = 0
        Remk1.RotX = 26
        sw4t.TimerEnabled = 1
    End If
End Sub

Sub sw4t_Timer
    Select Case sw4Step
        Case 1:Remk1.RotX = 14
        Case 2:Remk1.RotX = 2
        Case 3:Remk1.RotX = -10:sw4t.TimerEnabled = 0
    End Select

    sw4Step = sw4Step + 1
End Sub

Sub SolBallRelease(Enabled)
    If Enabled Then
        bsTrough.ExitSol_On
    End If
End Sub

Sub SolOutHole(Enabled)
    If Enabled Then
        bsTrough.EntrySol_On
        Drain.DestroyBall
    End If
End Sub

Sub SolDiv(Enabled)
    If Enabled Then
        PlaySound "fx_solenoidon"
        MGate.RotateToEnd
    Else
        PlaySound "fx_solenoidoff"
        MGate.RotateToStart
    End If
End Sub

'*****************
' Lock
'*****************

Sub lock3_Hit:Playsound "fx_sensor":Controller.Switch(39) = 1:End Sub

Sub SolBallLock(Enabled)
    If Enabled Then
        Controller.Switch(39) = 0
        lock1a.IsDropped = 1:lock2a.IsDropped = 1:lock3a.IsDropped = 1
        PlaySound SoundFX("fx_popper",DOFContactors)
        lock1.kick 180, 4:lock2.kick 180, 4:lock3.kick 180, 4
    Else
        lock1a.IsDropped = 0:lock2a.IsDropped = 0:lock3a.IsDropped = 0
    End If
End Sub

'**************
' Flipper Subs
'**************

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Dim LeftUP, RightUP
LeftUP = 0
RightUP = 0

Sub SolLFlipper(Enabled)
    If Enabled Then
        LeftUP = 1:CheckFlippers
        PlaySound SoundFX("fx_flipperup", DOFFlippers), 0, 1, -0.1, 0.15
        LeftFlipper.RotateToEnd
    Else
        LeftUP = 0
        PlaySound SoundFX("fx_flipperdown", DOFFlippers), 0, 1, -0.1, 0.15
        LeftFlipper.RotateToStart
    End If
End Sub

Sub SolRFlipper(Enabled)
    If Enabled Then
        RightUP = 1:CheckFlippers
        PlaySound SoundFX("fx_flipperup", DOFFlippers), 0, 1, 0.1, 0.15
        RightFlipper.RotateToEnd
    Else
        RightUP = 0
        PlaySound SoundFX("fx_flipperdown", DOFFlippers), 0, 1, 0.1, 0.15
        RightFlipper.RotateToStart
    End If
End Sub

Sub CheckFlippers
    If LeftUP AND RightUP Then
        vpmTimer.PulseSw 10
    End If
End Sub

Sub LeftFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 10, -0.1, 0.15
End Sub

Sub Rightflipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 10, 0.1, 0.15
End Sub

'******************
' RealTime Updates
'******************

Set MotorCallback = GetRef("RealTimeUpdates")

Sub RealTimeUpdates
	MgateP.RotZ = Mgate.CurrentAngle
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
OldGiState = -1 'start witht he Gi off

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

'******************************
' Diverse Collection Hit Sounds
'******************************

Sub aMetals_Hit(idx):PlaySound "fx_metalhit", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aRubber_Bands_Hit(idx):PlaySound "fx_rubber_band", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aRubber_Posts_Hit(idx):PlaySound "fx_rubber", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aRubber_Pins_Hit(idx):PlaySound "fx_postrubber", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aPlastics_Hit(idx):PlaySound "fx_PlasticHit", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aGates_Hit(idx):PlaySound "fx_Gate", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aWoods_Hit(idx):PlaySound "fx_Woodhit", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub

'********************
'Stern Free Fall
'added by Inkochnito
'********************
Sub editDips
    Dim vpmDips:Set vpmDips = New cvpmDips
    With vpmDips
        .AddForm 500, 400, "Free Fall - DIP switches"
        .AddFrame 0, 0, 190, "Maximum credits", &H00060000, Array("10 credits", 0, "15 credits", &H00020000, "25 credits", &H00040000, "40 credits", &H00060000)      'dip 18&19
        .AddFrame 0, 76, 190, "High game to date", 49152, Array("points", 0, "1 free game", &H00004000, "2 free games", 32768, "3 free games", 49152)                 'dip 15&16
        .AddFrame 0, 152, 190, "Special award", &HC0000000, Array("no award", 0, "100.000 points", &H40000000, "extra ball", &H80000000, "replay", &HC0000000)        'dip 31&32
        .AddFrame 0, 228, 190, "Add-a-ball memory", &H00801000, Array("1 ball only", 0, "3 balls", &H00800000, "5 balls", &H00801000)                                 'dip 13&24
        .AddChk 0, 300, 190, Array("Bonus multiplier in memory", &H10000000)                                                                                          'dip 29
        .AddChk 0, 315, 190, Array("Match feature", &H00100000)                                                                                                       'dip 21
        .AddChk 0, 330, 190, Array("Credits displayed", &H00080000)                                                                                                   'dip 20
        .AddFrame 205, 0, 190, "High score feature", &H00000020, Array("extra ball", 0, "replay", &H00000020)                                                         'dip 6
        .AddFrame 205, 48, 190, "Outlane special lites when", &H00400000, Array("completed card 2 times", 0, "completed card 1 time", &H00400000)                     'dip 23
        .AddFrame 205, 96, 190, "Outlane special lites when", &H00000010, Array("3 ball feature completed 2 times", 0, "3 ball feature completed 1 time", &H00000010) 'dip 5
        .AddFrame 205, 146, 190, "Arrow-card selector", &H00200000, Array("1 arrow on", 0, "2 arrows on", &H00200000)                                                 'dip 22
        .AddFrame 205, 193, 190, "Balls per game", &H00000040, Array("3 balls", 0, "5 balls", &H00000040)                                                             'dip 7
        .AddFrame 205, 242, 190, "Special limit", &H20000000, Array("1 per game", 0, "1 per ball", &H20000000)                                                        'dip 30
        .AddChk 205, 300, 190, Array("Background sound", &H00000080)                                                                                                  'dip 8
        .AddChk 205, 315, 190, Array("Talking feature", &H00008000)                                                                                                   'dip 17
        .AddChk 205, 330, 190, Array("Sky diver lites in memory", &H00002000)                                                                                         'dip 14
        .AddLabel 50, 350, 300, 15, "After hitting OK, press F3 to reset game with new settings."
        .ViewDips
    End With
End Sub
Set vpmShowDips = GetRef("editDips")

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


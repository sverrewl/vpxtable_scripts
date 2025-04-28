' ****************************
' MONOPOLY -  Stern 2001
' ****************************
'
'Release Notes
' 1.2  20230623 - VPX 10.7.3
'      - New rotating flipper implementation (single flipper and magnet) intended to correct ball being pushed
'        through walls, ramps and playfield
'      - Modified upper right flipper parameters to reduce interactions with rotating flipper
'      - Adjusted walls around the rotating flipper

Option Explicit
Randomize

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Const BallSize = 50

Const cGameName="monopoly",UseSolenoids=2,UseLamps=0,UseGI=0,SSolenoidOn="SolOn",SSolenoidOff="SolOff", SCoin="coin"
Const UseSync = 0
Const HandleMech = 0

LoadVPM "01120100", "sega.vbs", 3.23

Dim DesktopMode: DesktopMode = Table1.ShowDT
If DesktopMode = True Then 'Show Desktop components
Ramp16.visible=1
Ramp15.visible=1
Primitive13.visible=1
Else
Ramp16.visible=0
Ramp15.visible=0
Primitive13.visible=0
End if

'*************************************************************
'Solenoid Call backs
'**********************************************************************************************************
SolCallback(1) = "bsTrough.SolOut"
SolCallback(2) = "Autofire"
SolCallback(6) = "BankClose"
SolCallback(7) = "dtLL.SolDropUp" 'Drop Targets
SolCallback(8) = "LockKickBack"
SolCallback(12) = "bsChance.SolOut"
SolCallback(13) = "BankOpen"
SolCallback(19) = "SetLamp 119,"
SolCallback(20) = "SetLamp 120,"
SolCallback(21) = "SetLamp 121,"
SolCallback(22) = "SetLamp 122,"
SolCallback(23) = "SetLamp 123,"
SolCallback(25) = "WWMotor"             ' WaterWorks Motor (Handled by Mech Handler)
SolCallback(26) = "bsSaucerEC.SolOut"
SolCallback(27) = "MRelay"              ' Motor Relay
SolCallback(28) = "bsSaucerDE.SolOut"
SolCallback(29) = "SetLamp 129,"
SolCallback(30) = "DivertLeft"
SolCallback(31) = "DivertRight"
SolCallback(32) = "UpDown"

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
     If Enabled Then
         PlaySoundAtVol SoundFX("fx_Flipperup",DOFContactors), LeftFlipper, 1:LeftFlipper.RotateToEnd
     Else
         PlaySoundAtVol SoundFX("fx_Flipperdown",DOFContactors), LeftFlipper, 1:LeftFlipper.RotateToStart
     End If
  End Sub

Sub SolRFlipper(Enabled)
     If Enabled Then
         PlaySoundAtVol SoundFX("fx_Flipperup",DOFContactors), RightFlipper, 1:RightFlipper.RotateToEnd:RightFlipper1.RotateToEnd
     Else
         PlaySoundAtVol SoundFX("fx_Flipperdown",DOFContactors), RightFlipper, 1:RightFlipper.RotateToStart:RightFlipper1.RotateToStart
     End If
End Sub
'**********************************************************************************************************

'Solenoid Controlled toys
'**********************************************************************************************************

Sub Autofire(enabled)
    If enabled Then
        PlungerIM.AutoFire
    End If
End Sub

Sub BankClose(enabled)
    If enabled Then
        BankFlipper.RotateToStart
    PlaySoundAtVol SoundFX("fx_Flipperup",DOFContactors), BankFlipper, 1
    End If
End Sub

Sub BankOpen(enabled)
    If enabled Then
        BankFlipper.RotateToEnd
    PlaySoundAtVol SoundFX("fx_Flipperdown",DOFContactors), BankFlipper, 1
    End If
End Sub

Sub DivertLeft(enabled)
    If enabled Then
        DivLeft.RotatetoEnd
    PlaySoundAtVol SoundFX("fx_Flipperup",DOFContactors), DivLeft, 1
    Else
        DivLeft.RotateToStart
    End If
End Sub

Sub DivertRight(enabled)
    If enabled Then
        DivRight.RotatetoEnd
    PlaySoundAtVol SoundFX("fx_Flipperup",DOFContactors), DivRight, 1
    Else
        DivRight.RotateToStart
    End If
End Sub

Sub UpDown(enabled)
    If enabled Then
        UpDnPost.IsDropped = False
    PlaySoundAtVol SoundFX("Diverter",DOFContactors), UpDnPost, 1
    Else
        UpDnPost.IsDropped = True
    End If
End Sub

set GICallback = GetRef("UpdateGI")
Sub UpdateGI(no, Enabled)
    SetLamp 134, Enabled  'Backwall bulbs
  If Enabled Then
    dim xx
For each xx in Backlights:xx.visible = 1: Next
    For each xx in GI:xx.State = 1: Next
        PlaySound "fx_relay"

  Else
For each xx in Backlights:xx.visible = 0: Next
For each xx in GI:xx.State = 0: Next
        PlaySound "fx_relay"

  End If
End Sub

'**********************************************************************************************************

'Initiate Table
'**********************************************************************************************************

Dim bsTrough, bsChance, bsSaucerWW, bsSaucerEC, bsSaucerDE, WaterMech, BankMech, dtLL

Dim mMagnet

Sub Table1_Init
  vpmInit Me
  On Error Resume Next
    With Controller
    .GameName = cGameName
    If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
    .SplashInfoLine = "Monopoly Stern "&chr(13)&"You Suck"
    .HandleMechanics=0
    .HandleKeyboard=0
    .ShowDMDOnly=1
    .ShowFrame=0
    .ShowTitle=0
    .hidden = 0
         On Error Resume Next
         .Run GetPlayerHWnd
         If Err Then MsgBox Err.Description
         On Error Goto 0
     End With
     On Error Goto 0

    PinMAMETimer.Interval = PinMAMEInterval
    PinMAMETimer.Enabled = 1
    vpmNudge.TiltSwitch = swTilt
    vpmNudge.Sensitivity = 2
    vpmNudge.TiltObj=Array(Bumper1,Bumper2,Bumper3,Bumper4,Bumper5,Bumper6,LeftSlingshot,RightSlingshot)

    Set bsTrough = New cvpmBallStack
    With bsTrough
        .Initsw 0, 13, 12, 11, 14, 0, 0, 0
        .InitKick BallRelease, 90, 4
        .InitAddSnd "ballrelease"
        .InitEntrySnd "Drain", "Drain"
        .InitExitSnd SoundFX("ballrelease",DOFContactors), SoundFX("Solenoid",DOFContactors)
        .CreateEvents "bsTrough", Drain
        .IsTrough = True
        .Balls = 4
    End With

    ' Chance Scoop
    Set bsChance = New cvpmBallStack
    With bsChance
        .Initsw 0, 9, 0, 0, 0, 0, 0, 0
        .InitKick sw9, 120, 7
        .InitAddSnd "ballrelease"
        .InitExitSnd SoundFX("ballrelease",DOFContactors), SoundFX("Solenoid",DOFContactors)
        .InitEntrySnd "popper_ball", "solenoid"
        .CreateEvents "bsChance", sw9
    End With

    ' Saucer (Electric Company)
    Set bsSaucerEC = New cvpmBallStack
    With bsSaucerEC
        .InitSaucer sw26, 26, 315, 8
        .InitExitSnd SoundFX("Popper",DOFContactors), SoundFX("Solenoid",DOFContactors)
        .CreateEvents "bsSaucerEC", sw26
    End With

    ' Saucer (Dice Eject)
    Set bsSaucerDE = New cvpmBallStack
    With bsSaucerDE
        .InitSaucer sw52, 52, 90, 7
        .InitExitSnd SoundFX("Popper",DOFContactors), SoundFX("Solenoid",DOFContactors)
        .CreateEvents "bsSaucerDE", sw52
    End With

  'Drop Target
  Set dtLL=New cvpmDropTarget
    dtLL.InitDrop sw38,38
    dtLL.InitSnd SoundFX("DTDrop",DOFContactors),SoundFX("DTReset",DOFContactors)

    ' Saucer (Water Works)
    Set bsSaucerWW = New cvpmBallStack
    With bsSaucerWW
        .InitSaucer kick_29, 29, 180, 15
        .InitAltKick 340, 18
        .KickForceVar = 3
        .CreateEvents "bsSaucerWW", kick_29
    End With

  ' Magnet  (Water Works)
  Set mMagnet = New cvpmMagnet
    With mMagnet
      .InitMagnet magnetTrigger, 3
      .CreateEvents "mMagnet"
      .GrabCenter = False
    End With
    mMagnet.MagnetOn=True

    controller.switch(30) = True
    controller.switch(33) = True ' sw33-36 are optos and not handled correctly in current VPM (this corrects that)
    controller.switch(34) = True
    controller.switch(35) = True
    controller.switch(36) = True

    UpDnPost.IsDropped = True


End Sub

'**********************************************************************************************************
'Plunger code
'**********************************************************************************************************

Sub Table1_KeyDown(ByVal KeyCode)

  If keycode = PlungerKey Then Plunger.Pullback:PlaySoundAtVol "plungerpull", Plunger, 1
  If KeyDownHandler(keycode) Then Exit Sub
End Sub

Sub Table1_KeyUp(ByVal KeyCode)

  If keycode = PlungerKey Then Plunger.Fire:PlaySoundAtVol "plunger", Plunger, 1
  If KeyUpHandler(keycode) Then Exit Sub
End Sub

Dim plungerIM
    Const IMPowerSetting = 42 ' Plunger Power
    Const IMTime = 0.6        ' Time in seconds for Full Plunge
    Set plungerIM = New cvpmImpulseP
    With plungerIM
        .InitImpulseP swplunger, IMPowerSetting, IMTime
        .Random 0.3
        .Switch 16
        .InitExitSnd SoundFX("Popper",DOFContactors), SoundFX("Solenoid",DOFContactors)
        .CreateEvents "plungerIM"
    End With

'**********************************************************************************************************

 'kickers
Sub sw26_Hit:bsSaucerEC.AddBall 0 : PlaySoundAtVol "popper_ball", ActiveBall, 1: End Sub
Sub sw52_Hit:bsSaucerDE.AddBall 0 : PlaySoundAtVol "popper_ball", ActiveBall, 1: End Sub

'Watter Works Trigger
Sub sw29_Hit:Controller.Switch(29) = 1 : PlaySoundAtVol"rollover" , ActiveBall, 1: End Sub
Sub sw29_unHit:Controller.Switch(29) = 0:End Sub

'Ramp Gates
Sub sw10_hit:vpmTimer.PulseSw 10:End Sub
Sub sw39_hit:vpmTimer.PulseSw 39:End Sub
Sub sw48_hit:vpmTimer.PulseSw 48:End Sub
Sub sw31_hit:vpmTimer.PulseSw 31:End Sub

'Wire Triggers
Sub sw15_Hit:Controller.Switch(15) = 1 : PlaySoundAtVol"rollover" , ActiveBall, 1: End Sub
Sub sw15_unHit:Controller.Switch(15) = 0:End Sub
Sub sw16_Hit:: PlaySoundAtVol"rollover" , ActiveBall, 1: End Sub  'coded to impulse plunger
'Sub sw16_unHit:Controller.Switch(16) = 0:End Sub
Sub sw17_Hit:Controller.Switch(17) = 1 : PlaySoundAtVol"rollover" , ActiveBall, 1: End Sub
Sub sw17_unHit:Controller.Switch(17) = 0:End Sub
Sub sw18_Hit:Controller.Switch(18) = 1 : PlaySoundAtVol"rollover" , ActiveBall, 1: End Sub
Sub sw18_unHit:Controller.Switch(18) = 0:End Sub
Sub sw19_Hit:Controller.Switch(19) = 1 : PlaySoundAtVol"rollover" , ActiveBall, 1: End Sub
Sub sw19_unHit:Controller.Switch(19) = 0:End Sub
Sub sw25_Hit:Controller.Switch(25) = 1 : PlaySoundAtVol"rollover" , ActiveBall, 1: End Sub
Sub sw25_unHit:Controller.Switch(25) = 0:End Sub
Sub sw27_Hit:Controller.Switch(27) = 1 : PlaySoundAtVol"rollover" , ActiveBall, 1: End Sub
Sub sw27_unHit:Controller.Switch(27) = 0:End Sub
Sub sw28_Hit:Controller.Switch(28) = 1 : PlaySoundAtVol"rollover" , ActiveBall, 1: End Sub
Sub sw28_unHit:Controller.Switch(28) = 0:End Sub
Sub sw57_Hit:Controller.Switch(57) = 1 : PlaySoundAtVol"rollover" , ActiveBall, 1: End Sub
Sub sw57_unHit:Controller.Switch(57) = 0:End Sub
Sub sw58_Hit:Controller.Switch(58) = 1 : PlaySoundAtVol"rollover" , ActiveBall, 1: End Sub
Sub sw58_unHit:Controller.Switch(58) = 0:End Sub
Sub sw60_Hit:Controller.Switch(60) = 1 : PlaySoundAtVol"rollover" , ActiveBall, 1: End Sub
Sub sw60_unHit:Controller.Switch(60) = 0:End Sub
Sub sw61_Hit:Controller.Switch(61) = 1 : PlaySoundAtVol"rollover" , ActiveBall, 1: End Sub
Sub sw61_unHit:Controller.Switch(61) = 0:End Sub

'magnets optos
Sub sw33_Hit:Controller.Switch(33) = False:End Sub
Sub sw33_UnHit:Controller.Switch(33) = True:End Sub
Sub sw34_Hit:Controller.Switch(34) = False:End Sub
Sub sw34_UnHit:Controller.Switch(34) = True:End Sub
Sub sw35_Hit:Controller.Switch(35) = False:End Sub
Sub sw35_UnHit:Controller.Switch(35) = True:End Sub
Sub sw36_Hit:Controller.Switch(36) = False:End Sub
Sub sw36_UnHit:Controller.Switch(36) = True:End Sub

'Drop Targets
Sub Sw38_Dropped:dtLL.Hit 1 :End Sub

'Bumpers
Sub Bumper1_Hit : vpmTimer.PulseSw(41) : PlaySoundAtVol SoundFX("fx_bumper2",DOFContactors), ActiveBall, 1: End Sub
Sub Bumper2_Hit : vpmTimer.PulseSw(42) : PlaySoundAtVol SoundFX("fx_bumper2",DOFContactors), ActiveBall, 1: End Sub
Sub Bumper3_Hit : vpmTimer.PulseSw(43) : PlaySoundAtVol SoundFX("fx_bumper2",DOFContactors), ActiveBall, 1: End Sub

Sub Bumper4_Hit : vpmTimer.PulseSw(49) : PlaySoundAtVol SoundFX("fx_bumper1",DOFContactors), ActiveBall, 1: End Sub
Sub Bumper5_Hit : vpmTimer.PulseSw(50) : PlaySoundAtVol SoundFX("fx_bumper1",DOFContactors), ActiveBall, 1: End Sub
Sub Bumper6_Hit : vpmTimer.PulseSw(51) : PlaySoundAtVol SoundFX("fx_bumper1",DOFContactors), ActiveBall, 1: End Sub

 'Stand Up Targets
Sub sw44_Hit:vpmTimer.PulseSw 44:End Sub
Sub sw47_Hit:vpmTimer.PulseSw 47:End Sub
Sub sw46_Hit:vpmTimer.PulseSw 46:End Sub

'Spinners
Sub sw45_Spin:vpmTimer.PulseSw 45 : PlaySoundAtVol"fx_spinner" , sw45, 1: End Sub

'Generic Sounds
Sub Kicker1_Hit:PlaySoundAtVol "fx_ballrampdrop", ActiveBall, 1:End Sub
Sub Kicker2_Hit:PlaySoundAtVol "fx_ballrampdrop", ActiveBall, 1:End Sub
Sub Kicker3_Hit:PlaySoundAtVol "fx_ballrampdrop", ActiveBall, 1:End Sub
Sub Kicker4_Hit:PlaySoundAtVol "fx_ballrampdrop", ActiveBall, 1:End Sub
Sub Trigger5_Hit:PlaySoundAtVol "fx_ballrampdrop", ActiveBall, 1:End Sub

Sub Trigger1_Hit:PlaySoundAtVol "Wire Ramp", ActiveBall, 1:End Sub
Sub Trigger2_Hit:StopSound "Wire Ramp":PlaySoundAtVol "fx_ballrampdrop", ActiveBall, 1:End Sub


'************
'  3 Locks
'************
Sub Lock1_Hit() 'lower lock
    Controller.Switch(22) = 1
    Lock2.Enabled = 1
End Sub

Sub Lock2_Hit() 'middle lock
    Controller.Switch(21) = 1
    Lock3.Enabled = 1
End Sub

Sub Lock3_Hit() 'top lock
    Controller.Switch(20) = 1
End Sub

Sub LockKickBack(enabled)
    If Enabled then
        PlaySoundAtVol SoundFX("Popper",DOFContactors), Lock3, 1
        ' release top lock
        Lock3.enabled = False
        Lock3.Kick 0, 32
        Controller.Switch(20) = 0
        ' release middle lock
        Lock2.enabled = False
        Lock2.Kick 0, 32
        Controller.Switch(21) = 0
        ' release lower lock
        Lock1.Kick 0, 32
        Controller.Switch(22) = 0
    End If
End Sub

'*****************************
' WaterWorks Flipper
' - parts from Pacdude's pre-VPX table
' - modified to use a single flipper;
'   the multiple flipper approach had
'   balls going through walls, playfield and flipper.
'   the single flipper still passes through the ball occasionally
'*****************************
'Motor relay - direction
dim Direction:Direction = 0

Sub MRelay(enabled)
    Direction = abs(enabled)
End Sub

' WaterWorks Motor
Sub WWMotor(enabled)
    If enabled then
        UpdateFlipper.enabled = 1
    else
        UpdateFlipper.enabled = 0
    End If
End Sub

' WaterWorks FlipperController
dim WWPos, WWLastPos, WWSpeed, WWCounter, WWInterval:WWCounter = 0:WWInterval = 24
WWPos = 0:WWLastPos = WWPos

Sub UpdateFlipper_Timer()
    WWSpeed = abs(controller.getmech(1) ) ' Get speed from internal mech since only it has access to that value

    ' Speed select
    If WWSpeed = 0 Then UpdateFlipper.Interval = 28
    If WWSpeed = 1 Then UpdateFlipper.Interval = 26
    If WWSpeed = 2 Then UpdateFlipper.Interval = 14
    If WWSpeed = 3 Then UpdateFlipper.Interval = 12
    If WWSpeed = 4 Then UpdateFlipper.Interval = 7

    If Direction = 1 Then     ' Position Counter (direction flag comes from solenoid 27)
        WWPos = WWPos + 1
    Else
        WWPos = WWPos - 1
    End If
    If WWPos > 71 Then WWPos = 0
    If WWPos < 0 Then WWPos = 71

    ' ROM Service Menu and some parts of the manual identify switch 40 as the home position switch (not 30)
    ' and with switch 40 there are more directional changes
    If WWPos >= 0 and WWPos <= 3 Then controller.switch(40) = 1:else:controller.switch(40) = 0:end if 'home position
'    If WWPos >= 0 and WWPos <= 3 Then controller.switch(30) = 1:else:controller.switch(30) = 0:end if 'home position

    WWFlipper.Visible = False

    ' The Start & End angle setting (and this If statement) is needed to stop visual flipper jump that occurs
    ' when rotating clockwise and moving from position 71 to 0 (355 to 0 degrees)
    If Direction = 1 Then
        WWFlipper.StartAngle = INT(WWLastPos * 5)
        WWFlipper.EndAngle = INT(WWPos * 5)
    Else
        If NOT(WWPos=71 and WWLastPos=0) Then
            WWFlipper.StartAngle = INT(WWPos * 5)
            WWFlipper.EndAngle = INT(WWLastPos * 5)
        End If
    End If
    WWFlipper.RotateToEnd

    If WWPos >= 9 and WWPos <= 17 Then:kick_29.enabled = False:else:kick_29.enabled = True    ' saucer disabled while flipper moves over it
    If (WWPos = 8 or WWPos = 9) and Direction = 1 Then:bsSaucerWW.SolOut 1:End If            ' kick ball out at points near flipper movement
    If (WWPos = 20 or WWPos = 19) and Direction <> 1 Then:bsSaucerWW.SolOutAlt 1:End If

    WWLastPos = WWPos
    WWFlipper.Visible = True

End Sub

' Introduced magnet to help pull ball into kicker; flipper not reliable
Sub sw29_Hit
    mMagnet.MagnetOn=False
End Sub

Sub sw29_UnHit
    MagnetCycle.enabled=True
End Sub

Sub MagnetCycle_Timer
    me.enabled=False
    mMagnet.MagnetOn=True
End Sub


'Ball in Wall correction  move ball to Trigger 4
Sub Trigger3_Hit
    ActiveBall.X = 819.5076
  ActiveBall.Y = 1170.248
  ActiveBall.Z = 0
  ActiveBall.VelX = 0
  ActiveBall.VelY = 0
End Sub

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
    NFadeL 1, l1
    NFadeL 2, l2
    NFadeL 3, l3
    NFadeL 4, l4
    NFadeL 5, l5
    NFadeL 6, l6
    NFadeL 7, l7
    NFadeL 8, l8
    NFadeL 9, l9
    NFadeL 10, l10
    NFadeL 11, l11
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
    NfadeL 32, l32
    NFadeL 33, l33
    NFadeL 34, l34
    NFadeL 35, l35
    NFadeL 36, l36
    NFadeL 37, l37
    NFadeL 38, l38
    NFadeL 39, l39
    NFadeL 40, l40
    NFadeLm 41, l41a 'Bumper1
    NFadeL 41, l41
    NFadeLm 42, l42a 'Bumper3
    NFadeL 42, l42
'    NFadeObjm 43, P43, "bulbcover1_redOn", "bulbcover1_red"
 '   Flash 43, l43 'JackPot LED
FadeDisableLighting 43,p43,5
'    NFadeObjm 44, P44, "bulbcover1_blueOn", "bulbcover1_blue"
'    Flash 44, l44 'RailRoad LED
FadeDisableLighting 44,p44,5
    NFadeL 45, l45
    NFadeLm 46, l46a 'Bumper2
    NFadeL 46, l46
'    NFadeObjm 47, P47, "bulbcover1_yellowOn", "bulbcover1_yellow"
    NFadeLm 47, l47
'    Flash 47, l47a 'Community LED
FadeDisableLighting 47,p47,5
'    NFadeObjm 48, P48, "bulbcover1_yellowOn", "bulbcover1_yellow"
'    Flash 48, l48 'Free Parking LED
FadeDisableLighting 48,p48,5
FadeDisableLighting 49,p49,5
'    NFadeObjm 49, P49, "bulbcover1_orangeOn", "bulbcover1_orange"
    NFadeLm 49, l49
'    Flash 49, l49a 'Chance LED
    NFadeL 50, l50
    NFadeL 51, l51
    NFadeL 52, l52
    NFadeL 53, l53
    NFadeL 54, l54
    NFadeL 55, l55
    NFadeL 56, l56
    NFadeL 57, l57
    NFadeL 58, l58
    NFadeL 59, l59
    NFadeLm 60, l60a 'Bumper4
    NFadeL 60, l60
    NFadeLm 61, l61a 'Bumper5
    NFadeL 61, l61
    NFadeLm 62, l62a 'Bumper6
    NFadeL 62, l62
'    NFadeObjm 63, P63, "bulbcover1_greenOn", "bulbcover1_green"
'    Flash 63, l63 'Extra Ball LED
FadeDisableLighting 63,p64,5
'    NFadeObjm 64, P64, "bulbcover1_greenOn", "bulbcover1_green"
'    Flash 64, l64 'Roll & Collect LED
FadeDisableLighting 64,p63,5
    NFadeL 65, l65
    NFadeL 66, l66
    NFadeL 67, l67
    NFadeL 68, l68
    NFadeL 69, l69
    NFadeL 70, l70
    NFadeL 71, l71
    NFadeL 72, l72
    NFadeL 73, l73
    NFadeL 74, l74
    NFadeL 75, l75
    NFadeL 76, l76
    NFadeL 77, l77
    NFadeL 78, l78
    NFadeL 79, l79
    NFadeL 80, l80

' NFadeObjm 119, p119, "dome2_0_redON", "dome2_0_red"
    NFadeLm 119, L119
NFadeL 119, L119a
' NFadeObjm 120, p120, "dome2_0_redON", "dome2_0_red"
    NFadeLm 120, L120
NFadeL 120, L120a
' NFadeObjm 121, p121, "dome2_0_redON", "dome2_0_red"
    NFadeLm 121, L121
NFadeLm 121,l121b   ' 5000 When Flashing (Top    Playfield)
NFadeL 121, L121a
' NFadeObjm 122, p122, "dome2_0_redON", "dome2_0_red"
    NFadeLm 122, L122
NFadeLm 122,l122b   ' 5000 When Flashing (Bottom Playfield)
NFadeL 122, L122a
  NFadeObjm 123, p123, "dome2_0_clearON", "dome2_0_clear"
    Flash 123, f123
NFadeLm 123, L123a
NFadeLm 123, L123b
  NFadeObjm 129, p129, "dome2_0_clearON", "dome2_0_clear"
    Flash 129, f129
NFadeLm 129, L129a
NFadeLm 129, L129b




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

Sub FadeDisableLighting(nr, a, alvl)
  Select Case FadingLevel(nr)
    Case 4
      a.UserValue = a.UserValue - 0.1
      If a.UserValue < 0 Then
        a.UserValue = 0
        FadingLevel(nr) = 0
      end If
      a.BlendDisableLighting = alvl * a.UserValue 'brightness
    Case 5
      a.UserValue = a.UserValue + 0.50
      If a.UserValue > 1 Then
        a.UserValue = 1
        FadingLevel(nr) = 1
      end If
      a.BlendDisableLighting = alvl * a.UserValue 'brightness
  End Select
End Sub

'***************
'Mini DMD Array
'***************

Dim LED(15)

LED(0) = Array(d1, d2, d3, d4, d5, d6, d7)
LED(1) = Array(d8, d9, d10, d11, d12, d13, d14)
LED(2) = Array(d15, d16, d17, d18, d19, d20, d21)
LED(3) = Array(d22, d23, d24, d25, d26, d27, d28)
LED(4) = Array(d29, d30, d31, d32, d33, d34, d35)
LED(5) = Array(d36, d37, d38, d39, d40, d41, d42)
LED(6) = Array(d43, d44, d45, d46, d47, d48, d49)
LED(7) = Array(d50, d51, d52, d53, d54, d55, d56)
LED(8) = Array(d57, d58, d59, d60, d61, d62, d63)
LED(9) = Array(d64, d65, d66, d67, d68, d69, d70)
LED(10) = Array(d71, d72, d73, d74, d75, d76, d77)
LED(11) = Array(d78, d79, d80, d81, d82, d83, d84)
LED(12) = Array(d85, d86, d87, d88, d89, d90, d91)
LED(13) = Array(d92, d93, d94, d95, d96, d97, d98)
LED(14) = Array(d99, d100, d101, d102, d103, d104, d105)

Sub LedTimer_Timer
    Dim ChgLED, ii, num, chg, stat, obj, objf, x, y
    ChgLED = Controller.ChangedLEDs(&H00000000, &Hffffffff)
    If Not IsEmpty(ChgLED) Then
        For ii = 0 To UBound(chgLED)
            num = chgLED(ii, 0):chg = chgLED(ii, 1):stat = chgLED(ii, 2)
            For y = 0 to 6
                If chg And 1 Then
                    If(stat and 1) = 1 Then
                        Led(num) (y).color = RGB(255, 0, 0)
                    Else
                        led(num) (y).color = RGB(8, 8, 8)
                    End If
                End If
                chg = chg \ 4:stat = stat \ 4
            Next
        Next
    End If
End Sub




'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep

Sub RightSlingShot_Slingshot
  vpmTimer.PulseSw 62
    PlaySoundAtVol SoundFX("right_slingshot",DOFContactors), ActiveBall, 1
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.rotx = 20
    RStep = 0
    RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.rotx = 10
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.rotx = 0:RightSlingShot.TimerEnabled = 0:
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
  vpmTimer.PulseSw 59
    PlaySoundAtVol SoundFX("left_slingshot",DOFContactors), ActiveBall, 1
    LSling.Visible = 0
    LSling1.Visible = 1
    sling2.rotx = 20
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.rotx = 10
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.rotx = 0:LeftSlingShot.TimerEnabled = 0:
    End Select
    LStep = LStep + 1
End Sub

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

'********************************************************************
'      JP's VP10 Rolling Sounds (+rothbauerw's Dropping Sounds)
'********************************************************************

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

    For b = 0 to UBound(BOT)
        ' play the rolling sound for each ball
        If BallVel(BOT(b) ) > 1 AND BOT(b).z < 30 Then
            rolling(b) = True
            PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b)), AudioPan(BOT(b)), 0, Pitch(BOT(b)), 1, 0, AudioFade(BOT(b))
        Else
            If rolling(b) = True Then
                StopSound("fx_ballrolling" & b)
                rolling(b) = False
            End If
        End If

        ' play ball drop sounds
        If BOT(b).VelZ < -1 and BOT(b).z < 55 and BOT(b).z > 27 Then 'height adjust for ball drop sounds
            PlaySound "fx_ball_drop" & b, 0, ABS(BOT(b).velz)/17, AudioPan(BOT(b)), 0, Pitch(BOT(b)), 1, 0, AudioFade(BOT(b))
        End If
    Next
End Sub

'**********************
' Ball Collision Sound
'**********************

Sub OnBallBallCollision(ball1, ball2, velocity)
  PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 2000, AudioPan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub


'*****************************************
' ninuzzu's FLIPPER SHADOWS
'*****************************************

sub FlipperTimer_Timer()
  FlipperLSh.RotZ = LeftFlipper.currentangle
  FlipperRSh.RotZ = RightFlipper.currentangle
  FlipperRSh1.RotZ = RightFlipper1.currentangle

  Pr31.RotZ = sw31.currentangle
  PR39.RotZ = sw39.currentangle
  PR48.RotZ = sw48.currentangle

  BankDoor.RotY = - BankFlipper.CurrentAngle + 94

    RightSword1.RotZ = RightFlipper1.CurrentAngle
    RightSword.RotZ = RightFlipper.CurrentAngle
    LeftSword.RotZ = LeftFlipper.CurrentAngle

End Sub

'*****************************************
' BALL SHADOW
'*****************************************
Dim BallShadow
BallShadow = Array (BallShadow1,BallShadow2,BallShadow3,BallShadow4,BallShadow5)
Sub BallShadowUpdate_timer
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
    BallShadow(b).X = BOT(b).X
    ballShadow(b).Y = BOT(b).Y + 10
        If BOT(b).Z > 20 and BOT(b).Z < 140 Then
            BallShadow(b).visible = 1
        Else
            BallShadow(b).visible = 0
        End If
if BOT(b).z > 30 Then
ballShadow(b).height = BOT(b).Z - 20
ballShadow(b).opacity = 110
Else
ballShadow(b).height = BOT(b).Z - 24
ballShadow(b).opacity = 80
End If
    Next
End Sub



'************************************
' What you need to add to your table
'************************************

' a timer called RollingTimer. With a fast interval, like 10
' one collision sound, in this script is called fx_collide
' as many sound files as max number of balls, with names ending with 0, 1, 2, 3, etc
' for ex. as used in this script: fx_ballrolling0, fx_ballrolling1, fx_ballrolling2, fx_ballrolling3, etc


'******************************************
' Explanation of the rolling sound routine
'******************************************

' sounds are played based on the ball speed and position

' the routine checks first for deleted balls and stops the rolling sound.

' The For loop goes through all the balls on the table and checks for the ball speed and
' if the ball is on the table (height lower than 30) then then it plays the sound
' otherwise the sound is stopped, like when the ball has stopped or is on a ramp or flying.

' The sound is played using the VOL, AUDIOPAN, AUDIOFADE and PITCH functions, so the volume and pitch of the sound
' will change according to the ball speed, and the AUDIOPAN & AUDIOFADE functions will change the stereo position
' according to the position of the ball on the table.


'**************************************
' Explanation of the collision routine
'**************************************

' The collision is built in VP.
' You only need to add a Sub OnBallBallCollision(ball1, ball2, velocity) and when two balls collide they
' will call this routine. What you add in the sub is up to you. As an example is a simple Playsound with volume and paning
' depending of the speed of the collision.


Sub Pins_Hit (idx)
  PlaySound "pinhit_low", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Targets_Hit (idx)
  PlaySound "target", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Metals_Thin_Hit (idx)
  PlaySound "metalhit_thin", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Metals_Medium_Hit (idx)
  PlaySound "metalhit_medium", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Metals2_Hit (idx)
  PlaySound "metalhit2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Gates_Hit (idx)
  PlaySound "gate4", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Spinner_Spin
  PlaySound "fx_spinner", 0, .25, AudioPan(Spinner), 0.25, 0, 0, 1, AudioFade(Spinner)
End Sub

Sub Rubbers_Hit(idx)
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 20 then
    PlaySound "fx_rubber2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End if
  If finalspeed >= 6 AND finalspeed <= 20 then
    RandomSoundRubber()
  End If
End Sub

Sub Posts_Hit(idx)
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

' Thalamus : Exit in a clean and proper way
Sub Table1_exit
  Controller.Pause = False
  Controller.Stop
End Sub

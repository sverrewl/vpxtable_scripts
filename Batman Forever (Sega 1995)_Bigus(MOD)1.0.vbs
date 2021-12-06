'############################################################################################
'############################################################################################
'#######                                                                             ########
'#######          Batman Forever                                                     ########
'#######          (Sega 1995)                                                        ########
'#######                                                                             ########
'############################################################################################
'############################################################################################

'VP9 Version 1.2 Desktop mfuegemann 2014

' Thanks to:
' Fuzzel for providing the images for the playfield lights
' JPSalas for the Flasher fading code

'Crude VPX conversion by DevaL

Option Explicit

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Const cgamename = "batmanf"

'-----------------------------------
'------  Global Cofigurations ------
'-----------------------------------
Const DimGI=-25         'set between -100 and 105 to dim or brighten GI lights (minus is darker)
Const DimFlashers=0     'set between -255 and 0 to dim or brighten the Flashers (minus is darker)

' Thalamus 2020 March : Improved directional sounds
' !! NOTE : Table not verified yet !!

' Options
' Volume devided by - lower gets higher sound

Const VolDiv = 2000    ' Lower number, louder ballrolling/collition sound
Const VolCol = 10      ' Ball collition divider ( voldiv/volcol )

Const UseSolenoids=2,UseSync=1
Const SSolenoidOn="SolOn",SSolenoidOff="SolOff",SFlipperOn="FlipperUp",SFlipperOff="FlipperDown",SCoin="coin3"

LoadVPM "01560000","DE.VBS",3.1

'-----------------------------------
'------  Solenoid Assignment  ------
'-----------------------------------
SolCallback(1) = "SolTroughBall"                        'Trough Lock Ball
SolCallback(2) = "bsBallRelease.SolOut"                 'Trough Up-Kicker
SolCallback(3) = "Autoplunger"                          'Auto Launch
SolCallback(4) = "SolVUK"                               'VUK
SolCallback(5) = "SolTopEject"                          'Top Eject
'6 - not used
SolCallback(7) = "DropTargetBank.SolDropUp"             'Drop Target Reset
SolCallback(8) = "vpmSolSound SoundFX(""knocker"",DOFKnocker),"             'KKnocker
SolCallback(9) = "bsBallLock.SolOut"                    'Bat Cave Ball Lock
SolCallback(10) = "Sol10"                   'Left/Right AB Relay
SolCallback(11) = "SolGI"                   'GI
SolCallback(12) = "SolRampDiverter"                     'Batwing Ramp Diverter
'13 - not used
'SolCallback(14) 'Batwing Motor - handled elsewhere
SolCallback(15) = "SolFireBWGun"                        'Cannon Launch
SolCallback(16) = "SolBatCaveExit"                      'Bat Cave Exit Kicker
'SolCallback(17) = "vpmSolSound SoundFX(""Bumper"",DOFContactors),"             'Left Bumper
'SolCallback(18) = "vpmSolSound SoundFX(""Bumper"",DOFContactors),"             'Bottom Bumper
'SolCallback(19) = "vpmSolSound SoundFX(""Bumper"",DOFContactors),"             'Right Bumper
'SolCallback(20) = "vpmSolSound SoundFX(""lSling"",DOFContactors),"             'Left Slingshot
'SolCallback(21) = "vpmSolSound SoundFX(""rSling"",DOFContactors),"             'Right Slingshot
SolCallback(22) = "vpmSolAutoPlunger LaserKick,200,"    'Laser Kick
SolCallback(25) = "SolFlasher1R"
SolCallback(26) = "SolFlasher2R"
SolCallback(27) = "SolFlasher3R"
SolCallback(28) = "SolFlasher4R"
SolCallback(29) = "SolFlasher5R"
SolCallback(30) = "SolFlasher6R"
SolCallback(31) = "SolFlasher7R"
SolCallback(32) = "SolFlasher8R"
'SolCallback(sLRFlipper)="vpmSolFlipper RightFlipper,URightFlipper,"
'SolCallback(sLLFlipper)="vpmSolFlipper LeftFlipper,nothing,"

 SolCallback(sLLFlipper) = "SolLFlipper"
 SolCallback(sLRFlipper) = "SolRFlipper"

    Sub SolLFlipper(Enabled)
        If Enabled Then
            PlaySoundAtVol SoundFX("FlipperUp",DOFFlippers), LeftFlipper, 1
            LeftFlipper.RotateToEnd
        Else
            PlaySoundAtVol SoundFX("FlipperDown",DOFFlippers), LeftFlipper, 1
            LeftFlipper.RotateToStart
        End If
    End Sub

    Sub SolRFlipper(Enabled)
        If Enabled Then
            PlaySoundAtVol SoundFX("FlipperUp",DOFFlippers), RightFlipper, 1
            RightFlipper.RotateToEnd
            URightFlipper.RotateToEnd
        Else
            PlaySoundAtVol SoundFX("FlipperDown",DOFFlippers), RightFlipper, 1
            RightFlipper.RotateToStart
            URightFlipper.RotateToStart
        End If
    End Sub


Sub SolTroughBall(enabled)
    if enabled then
        bsTrough.ExitSol_On
        bsBallRelease.AddBall 0
        vpmCreateBall BallRelease
    end if
End Sub

Sub AutoPlunger(Enabled)
    if enabled then
        Plunger.Kick 0,35+(Rnd*0.9)     '34.4-33.5
        Playsound SoundFX("launch",DOFContactors)
    end if
End Sub

Dim DiverterUp
Sub SolRampDiverter(enabled)
    if enabled then
        RampDiverter.collidable = false
        DiverterUp = True
        DiverterTimer.enabled = True
    else
        RampDiverter.collidable = true
        DiverterUp = False
        DiverterTimer.enabled = True
    end if
end sub

Sub DiverterTimer_Timer
    if DiverterUp then
        pDiverter.Rotx = pDiverter.Rotx + 4
        if pDiverter.Rotx >= 50 then
            pDiverter.Rotx = 50
            DiverterTimer.enabled = False
        end if
    else
        pDiverter.Rotx = pDiverter.Rotx - 1
        if pDiverter.Rotx <= 3 then
            pDiverter.Rotx = 3
            DiverterTimer.enabled = False
        end if
    end if
End Sub

Dim obj,bsTrough,bsTroughKick,bsBallRelease,DropTargetBank,bsBallLock,mBWGun

'------------------------
'-----  Table Init  -----
'------------------------

Sub BF_Init
vpmInit Me
    vpminit me
    On Error Resume Next
        With Controller
        .GameName = cGameName
        If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
        .SplashInfoLine = "Batman Forever (Sega 1995)" & vbNewLine & "created by mfuegemann" & vbNewLine & "crude VPX conversion by DevaL"
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

    ' DMD position for 3 Monitor Setup
    'Controller.Games(cGameName).Settings.Value("dmd_pos_x")=3850
    'Controller.Games(cGameName).Settings.Value("dmd_pos_y")=300
    'Controller.Games(cGameName).Settings.Value("dmd_width")=505
    'Controller.Games(cGameName).Settings.Value("dmd_height")=155
    'Controller.Games(cGameName).Settings.Value("rol")=0

    PinMAMETimer.Interval = PinMAMEInterval
    PinMAMETimer.Enabled = 1

    vpmNudge.TiltSwitch=1
    vpmNudge.Sensitivity=2
    vpmNudge.TiltObj=Array(Bumper57,Bumper58,Bumper59,LeftSlingshot,RightSlingshot)

    vpmMapLights AllLights

    Set bsTrough = New cvpmBallStack
        bsTrough.InitSw 0,14,13,12,11,0,0,0
        bsTrough.CreateEvents "bsTrough", Drain
        bsTrough.InitAddSnd SSolenoidOn
        bsTrough.InitEntrySnd "drain5", SSolenoidOn
        bsTrough.Balls=4

    set bsBallRelease = New cvpmBallStack
        bsBallRelease.InitSaucer BallRelease, 15, 90, 5
        bsBallRelease.InitExitSnd SoundFX("ballrel",DOFContactors),SoundFX("solenoid",DOFContactors)

    set DropTargetBank = new cvpmDropTarget
        DropTargetBank.InitDrop Array(DropTargetBottom,DropTargetBotMid,DropTargetTopMid,DropTargetTop), Array(18,19,20,21)
        DropTargetBank.InitSnd SoundFX("Targetdrop1",DOFDropTargets),SoundFX("TargetBankreset1",DOFContactors)

    Set bsBallLock = New cvpmBallStack
        bsBallLock.InitSw 0,62,61,60,0,0,0,0
        bsBallLock.InitKick BallLock, 180, 5
        bsBallLock.CreateEvents "bsBallLock", BallLock
        bsBallLock.InitEntrySnd "solenoid", SSolenoidOn
        bsBallLock.Balls=0

    Set mBWGun = New cvpmMech
        mBWGun.MType = vpmMechOneSol + vpmMechReverse + vpmMechNonLinear
        mBWGun.Sol1 = 14
        mBWGun.Length = 500
        mBWGun.Steps = 200
        mBWGun.AddSw 46,0,3
        mBWGun.AddSw 46,197,200
        mBWGun.Callback = GetRef("UpdateBWGun")
        mBWGun.Start

    LaserKick.pullback

    RampDiverter.collidable = true
    DiverterUp = false
    pDiverter.Rotx = 50
    DiverterTimer.enabled = true

    if (DimGI >= -100) and (DimGI <= 105) then
        For each obj in GIString1
            'obj.Alpha = obj.Alpha + DimGI
            'obj.isvisible = True
        Next
        For each obj in GIString2
            'obj.Alpha = obj.Alpha + DimGI
            'obj.isvisible = True
        Next
    end if
End Sub

Sub BF_Exit
    Controller.Stop
End Sub

Sub TopEjectKicker_Hit
    Controller.Switch(41) = True
End Sub

Sub SolTopEject(Enabled)
    If Enabled Then
        If Controller.Switch(41) = True Then
            PlaySound SoundFX("Kicker",DOFContactors)
            TopEjectKicker.Kick 90,9
            Controller.Switch(41) = False
        end if
    end if
End Sub

Sub BatCaveExitKicker_Hit
    Controller.Switch(30) = True
End Sub

Sub SolBatCaveExit(Enabled)
    If Enabled Then
        If Controller.Switch(30) = True Then
            PlaySound SoundFX("Kicker",DOFContactors)
            BatCaveExitKicker.Kick 180,45
            Controller.Switch(30) = False
        end if
    end if
End Sub


'---------------------------
'------  GI Handling  ------
'---------------------------

Sub SolGI(enabled)
    if enabled then
        For each obj in GIString2
            obj.visible = False
            'obj.triggersingleupdate
        Next
    else
        For each obj in GIString2
            obj.visible = True
            'obj.triggersingleupdate
        Next
    end if
End Sub

Sub Sol10(enabled)
    if enabled then
        For each obj in GIString1
            obj.visible = True
            'obj.triggersingleupdate
        Next
    else
        For each obj in GIString1
            obj.visible = False
            'obj.triggersingleupdate
        Next
    end if
End Sub


'--------------------------------
'------  Flasher Handling  ------
'--------------------------------

Sub SolFlasher1R(enabled)
    if enabled then
        FLamp1R.state = 1
        Setflash 6, True
    else
        FLamp1R.state = 0
        Setflash 6, False
    end if
End Sub

Sub SolFlasher2R(enabled)
    if enabled then
        FLamp2Ra.state = 1
        FLamp2Rb.state = 1
        Setflash 4, True
    else
        FLamp2Ra.state = 0
        FLamp2Rb.state = 0
        Setflash 4, False
    end if
End Sub

Sub SolFlasher3R(enabled)
    if enabled then
     Flasher3R.state = 1
Setflash 5, True
    else
    Flasher3R.state = 0
Setflash 5, False
    end if
End Sub

Sub SolFlasher4R(enabled)
    if enabled then
        Flasher4R.state = 1
        Flasher4Ra.state = 1
Setflash 0, True
    else
        Flasher4R.state = 0
        Flasher4Ra.state = 0
Setflash 0, False
    end if
End Sub

Sub SolFlasher5R(enabled)
    if enabled then
      Flasher5R.state = 1
Setflash 1, True
    else
       Flasher5R.state = 0
Setflash 1, False
    end if
End Sub

Sub SolFlasher6R(enabled)
    if enabled then
     Flasher6R.state = 1
Setflash 2, True
    else
        Flasher6R.state = 0
Setflash 2, False
    end if
End Sub

Sub SolFlasher7R(enabled)
    if enabled then
       Flasher7R.state = 1
Setflash 3, True
    else
       Flasher7R.state = 0
Setflash 3, False
    end if
End Sub

Sub SolFlasher8R(enabled)
    if enabled then
        FLamp8R.state = 1
        Setflash 7, True
    else
        FLamp8R.state = 0
        Setflash 7, False
    end if
End Sub


'---------------------------
'------  Batwing Gun  ------
'---------------------------

Sub UpdateBWGun(aNewPos, aSpeed, aLastPos)
    if anewpos < 100 then
        pBatwing.rotz = -anewpos / 2
        pBatwingGun.rotz = anewpos / 2
        GunAngle = anewpos / 2
    else
        pBatwing.rotz = (anewpos - 200) / 2
        pBatwingGun.rotz = (200 - anewpos) / 2
        GunAngle = (200 - anewpos) / 2
    end if
End Sub

Sub BWGun_Hit
    controller.switch(44) = True
End Sub

Dim GunAngle
Sub SolFireBWGun(Enabled)
    if enabled then
        BWGun.kick GunAngle, 50
        controller.switch(44) = False
        vpmTimer.PulseSw 45
    end if
End Sub


'-----------------------------
'------  VUK animation  ------
'-----------------------------
Sub VUK_Bottom_Hit
    Controller.Switch(32) = True
End Sub

Sub SolVUK(Enabled)
    If Enabled Then
        If Controller.Switch(32) = True Then
            PlaySound SoundFX("Kicker",DOFContactors)
            VUK_Bottom.destroyball
            VUK1.CreateBall
            vpmTimer.AddTimer 70,"VUKLevel1"
        end if
    end if
end sub

Sub VUKLevel1(swNo)
    Controller.Switch(32) = False
    VUK1.DestroyBall
    VUK2.CreateBall
    vpmTimer.AddTimer 70,"VUKLevel2"
End Sub

Sub VUKLevel2(swNo)
    VUK2.DestroyBall
    VUK3.CreateBall
    vpmTimer.AddTimer 70,"VUKLevel3"
End Sub

Sub VUKLevel3(swNo)
    VUK3.DestroyBall
    VUK4.CreateBall
    vpmTimer.AddTimer 70,"VUKLevel4"
End Sub

Sub VUKLevel4(swNo)
    VUK4.DestroyBall
    VUK_Top.CreateBall
    VUK_Top.Kick 330,15
End Sub


'-------------------------------
'------  Keybord Handler  ------
'-------------------------------

Sub BF_KeyDown(ByVal keycode)
    If keycode = LeftFlipperKey  Then Controller.Switch(63) = True
    If keycode = RightFlipperKey Then Controller.Switch(64) = True
    If KeyDownHandler(keycode) Then Exit Sub
    If keycode=PlungerKey Then
        Controller.Switch(49)  = 1          'Auto Launch
    End if
End Sub

Sub BF_KeyUp(ByVal keycode)
    If keycode = LeftFlipperKey  Then Controller.Switch(63) = False
    If keycode = RightFlipperKey Then Controller.Switch(64) = False
    If KeyUpHandler(keycode) Then Exit Sub
    If keycode=PlungerKey Then
        Controller.Switch(49)  = 0          'Auto Launch
    End if
End Sub


'---------------------------------
'------  Switch Assignment  ------
'---------------------------------

'1  Tilt - handled elsewhere
'2  Coin4 - handled elsewhere
'3  Start - handled elsewhere
'4  Coin1 - handled elsewhere
'5  Coin2 - handled elsewhere
'6  Coin3 - handled elsewhere
'7  Slam Tilt - handled elsewhere
'8  Buy In - handled elsewhere
'9  not used
'10 not used
'11 Ball1 - handled elsewhere
'12 Ball2 - handled elsewhere
'13 Ball3 - handled elsewhere
'14 Ball4 - handled elsewhere
'15 Trough Opto - handled elsewhere
Sub Trigger16_Hit:Controller.switch(16) = True:end sub      '16 Shooter Lane
Sub Trigger16_Unhit:Controller.switch(16) = False:end sub
'17 not used
Sub DropTargetBottom_Hit:vpmTimer.pulseSwitch(18),0,"":End Sub      '18 L4Bank Bottom
Sub DropTargetBotMid_Hit:vpmTimer.pulseSwitch(19),0,"":End Sub      '19 L4Bank MidBottom
Sub DropTargetTopMid_Hit:vpmTimer.pulseSwitch(20),0,"":End Sub      '20 L4Bank TopBottom
Sub DropTargetTop_Hit:vpmTimer.pulseSwitch(21),0,"":End Sub         '21 L4Bank Top
Sub DropTargetBottom_Hit:DropTargetBank.Hit 1:DropTargetBottom.isdropped=true:End Sub      '18 and one for the Queen
Sub DropTargetBotMid_Hit:DropTargetBank.Hit 2:DropTargetBotMid.isdropped=true:End Sub      '19
Sub DropTargetTopMid_Hit:DropTargetBank.Hit 3:DropTargetTopMid.isdropped=true:End Sub      '20
Sub DropTargetTop_Hit:DropTargetBank.Hit 4:DropTargetTop.isdropped=true:End Sub            '21
Sub Trigger22_Hit:Controller.Switch(22)=1:End Sub      '22  Top Lane Left
Sub Trigger22_Unhit:Controller.Switch(22)=0:End Sub
Sub Trigger23_Hit:Controller.Switch(23)=1:End Sub      '23  Top Lane Middle
Sub Trigger23_Unhit:Controller.Switch(23)=0:End Sub
Sub Trigger24_Hit:Controller.Switch(24)=1:End Sub      '24  Top Lane Right
Sub Trigger24_Unhit:Controller.Switch(24)=0:End Sub
Sub TargetRampLeft_Hit:vpmTimer.PulseSw 25:PlaysoundAtVol SoundFX("Target1",DOFTargets), ActiveBall, 1:End Sub      '25 Left Ramp Left
Sub TargetRampRight_Hit:vpmTimer.PulseSw 26:PlaysoundAtVol SoundFX("Target1",DOFTargets), ActiveBall, 1:End Sub     '26 Left Ramp Right
Sub TargetVUK_Hit:vpmTimer.PulseSw 27:PlaysoundAtVol SoundFX("Target1",DOFTargets), ActiveBall, 1:End Sub       '27 VUK
Sub TargetRightRampLeft_Hit:vpmTimer.PulseSw 28:PlaysoundAtVol SoundFX("Target1",DOFTargets), ActiveBall, 1:End Sub     '28 Right Ramp Left Standup
Sub TargetRightRampRight_Hit:vpmTimer.PulseSw 29:PlaysoundAtVol SoundFX("Target1",DOFTargets), ActiveBall, 1:End Sub        '29 Right Ramp Right ?
'30 Bat Cave Exit Kicker - handled elsewhere
'31 not used
'32 VUK - handled elsewhere
Sub Trigger33_Hit                                      '33  Laser Kick Left Outlane
    vpmTimer.Pulsesw 33
    if Activeball.vely > 0 then
        Activeball.vely = 0.5 * Activeball.vely
    end if
End Sub
Sub Trigger33_Unhit:Controller.Switch(33)=0:End Sub
Sub Trigger34_Hit:Controller.Switch(34)=1:End Sub      '34  Left Return Lane
Sub Trigger34_Unhit:Controller.Switch(34)=0:End Sub
Sub LeftSlingshot_Slingshot()                           '35 Left Slingshot
    vpmTimer.pulseSwitch(35),0,""
    PlaySoundAtVol SoundFX("Bumper",DOFContactors), ActiveBall, 1
End Sub
Sub RightSlingshot_Slingshot()                          '36 Right Slingshot
    vpmTimer.pulseSwitch(36),0,""
    PlaySoundAtVol SoundFX("Bumper",DOFContactors), ActiveBall, 1
End Sub
Sub Trigger37_Hit:Controller.Switch(37)=1:End Sub      '37  Right Return Lane inside
Sub Trigger37_Unhit:Controller.Switch(37)=0:End Sub
Sub Trigger38_Hit:Controller.Switch(38)=1:End Sub      '38  Right Return Lane outside
Sub Trigger38_Unhit:Controller.Switch(38)=0:End Sub
Sub Trigger39_Hit:Controller.Switch(39)=1:End Sub      '39  Right Outlane
Sub Trigger39_Unhit:Controller.Switch(39)=0:End Sub
Sub Trigger40_Hit:Controller.Switch(40)=1:End Sub      '40  U-Turn Ramp
Sub Trigger40_Unhit:Controller.Switch(40)=0:End Sub
'41 Top Eject - handled elsewhere
Sub TargetTopGood_Hit:vpmTimer.PulseSw 42:PlaysoundAtVol SoundFX("Target1",DOFTargets), ActiveBall, 1:End Sub       '42 Top Good Head Standup
Sub TargetTopBad_Hit:vpmTimer.PulseSw 43:PlaysoundAtVol SoundFX("Target1",DOFTargets), ActiveBall, 1:End Sub        '43 Top Bad Head Standup
'44 Batwing Loaded - handled elsewhere
'45 Batwing Location - handled elsewhere
'46 Batwing Home - handled elsewhere
Sub Trigger47_Hit:Controller.Switch(47)=1:End Sub      '47  Left Orbit
Sub Trigger47_Unhit:Controller.Switch(47)=0:End Sub
Sub Trigger48_Hit:Controller.Switch(48)=1:End Sub      '48  Right Orbit
Sub Trigger48_Unhit:Controller.Switch(48)=0:End Sub
'49 Gun Trigger (Auto Launch) - handled elsewhere
Sub TargetRightGood_Hit:vpmTimer.PulseSw 50:PlaysoundAtVol SoundFX("Target1",DOFTargets), ActiveBall, 1:End Sub     '50 Right Good Head Standup
Sub TargetRightBad_Hit:vpmTimer.PulseSw 51:PlaysoundAtVol SoundFX("Target1",DOFTargets), ActiveBall, 1:End Sub      '51 Right Bad Head Standup
Sub Trigger52_Hit:Controller.Switch(52)=1:End Sub      '52  Left Ramp Enter
Sub Trigger52_Unhit:Controller.Switch(52)=0:End Sub
Sub Trigger53_Hit:Controller.Switch(53)=1:End Sub      '53  Left Ramp Exit - Manual is wrong
Sub Trigger53_Unhit:Controller.Switch(53)=0:End Sub
Sub Trigger54_Hit:Controller.Switch(54)=1:End Sub      '54  Middle Ramp Exit - Manual is wrong
Sub Trigger54_Unhit:Controller.Switch(54)=0:End Sub
Sub Trigger55_Hit:Controller.Switch(55)=1:End Sub      '55  Right Ramp Enter
Sub Trigger55_Unhit:Controller.Switch(55)=0:End Sub
Sub Trigger56_Hit:Controller.Switch(56)=1:End Sub      '56  Right Ramp Exit
Sub Trigger56_Unhit:Controller.Switch(56)=0:End Sub
Sub Bumper57_Hit()                          '57 Bumper Left
    vpmTimer.pulseSwitch(57),0,""
    PlaySoundAtVol SoundFX("Bumper",DOFContactors), ActiveBall, 1
    Bumper57_Dir = -1
    Bumper57.Timerenabled = False
    Bumper57.Timerenabled = True
End Sub
Sub Bumper58_Hit()                          '58 Bumper Bottom
    vpmTimer.pulseSwitch(58),0,""
    PlaySoundAtVol SoundFX("Bumper",DOFContactors), ActiveBall, 1
    Bumper58_Dir = -1
    Bumper58.Timerenabled = False
    Bumper58.Timerenabled = True
End Sub
Sub Bumper59_Hit()                          '59 Bumper Right
    vpmTimer.pulseSwitch(59),0,""
    PlaySoundAtVol SoundFX("Bumper",DOFContactors), ActiveBall, 1
    Bumper59_Dir = -1
    Bumper59.Timerenabled = False
    Bumper59.Timerenabled = True
End Sub
'60 Lock 3 Top - handled elsewhere
'61 Lock 2 Middle - handled elsewhere
'62 Lock 1 Bottom - handled elsewhere
'63 Left Flipper - handled elsewhere
'64 Upper+Lower Right Flipper - handled elsewhere

'--------------------------------------------
'Helper Functions
'--------------------------------------------

'Bumper animation
Dim Bumper57_Dir,Bumper58_Dir,Bumper59_Dir

Sub Bumper57_Timer
    P_BumperRing57.TransZ = P_BumperRing57.TransZ + Bumper57_Dir * 2.5
    if P_BumperRing57.TransZ <= -30 then Bumper57_Dir = 1
    if P_BumperRing57.TransZ >= 0 then
        Bumper57.Timerenabled = False
        Bumper57_Dir = 0
    end if
End Sub

Sub Bumper58_Timer
    P_BumperRing58.TransZ = P_BumperRing58.TransZ + Bumper58_Dir * 2.5
    if P_BumperRing58.TransZ <= -30 then Bumper58_Dir = 1
    if P_BumperRing58.TransZ >= 0 then
        Bumper58.Timerenabled = False
        Bumper58_Dir = 0
    end if
End Sub

Sub Bumper59_Timer
    P_BumperRing59.TransZ = P_BumperRing59.TransZ + Bumper59_Dir * 2.5
    if P_BumperRing59.TransZ <= -30 then Bumper59_Dir = 1
    if P_BumperRing59.TransZ >= 0 then
        Bumper59.Timerenabled = False
        Bumper59_Dir = 0
    end if
End Sub


Sub RampHelper1_Hit
    if Activeball.velz > 1 then Activeball.velz = 0
    if abs(Activeball.velx) > 35 then Activeball.velx = 35 * (Activeball.velx/abs(Activeball.velx))
    if abs(Activeball.vely) > 35 then Activeball.vely = 35 * (Activeball.vely/abs(Activeball.vely))
End Sub

Sub RampHelper2_Hit
    Activeball.velz = 0
End Sub

Sub RampHelper3_Hit
    Activeball.velz = 0
    if abs(Activeball.velx) > 15 then Activeball.velx = 15 * (Activeball.velx/abs(Activeball.velx))
    if abs(Activeball.vely) > 15 then Activeball.vely = 15 * (Activeball.vely/abs(Activeball.vely))
End Sub

Sub RampHelper4_Hit
    Activeball.velz = 0.15 * Activeball.velz
    if abs(Activeball.velx) > 50 then Activeball.velx = 50 * (Activeball.velx/abs(Activeball.velx))
    if abs(Activeball.vely) > 50 then Activeball.vely = 50 * (Activeball.vely/abs(Activeball.vely))
End Sub


'---------------------------------------
'------  JP's Flasher Fading Sub  ------
'---------------------------------------
Dim Flashers
Flashers = Array(Flasher4Ra,Flasher5Ra,Flasher6Ra,Flasher7R,Flasher2Ra,Flasher3R,Flasher1R,Flasher8R,Flasher6Rb,Flasher5Rb,Flasher2Rb,Flasher4Ra1,Flasher4Rb,Flasher4Rb1,Flasher5Rb1,Flasher6Rb1,Flasher7R1,Flasher3R1)

Dim FlashMaxAlpha
Dim FlashState(200), FlashLevel(200)
Dim FlashSpeedUp, FlashSpeedDown

FlashInit()
FlasherTimer.Interval = 5
FlasherTimer.Enabled = 1

Sub FlashInit
    Dim i
    For i = 0 to 200
        FlashState(i) = 0
        FlashLevel(i) = 0
    Next
    FlashSpeedUp = 50    'fast speed when turning on the flasher
    FlashSpeedDown = 10  'slow speed when turning off the flasher, gives a smooth fading
    ' you could also change the default images for each flasher or leave it as in the editor
    ' for example
    ' Flasher1.Image = "fr"

    'added by mfuegemann to apply Dim settings
    FlashMaxAlpha = 255
'   if DimFlashers < 0 then
'       FlashMaxAlpha = FlashMaxAlpha + DimFlashers
'       if FlashMaxAlpha < 0 then
'           FlashMaxAlpha = 0
'       end if
'   end if

    AllFlashOff()
End Sub

Sub AllFlashOff
    Dim i
    For i = 0 to 200
        FlashState(i) = 0
    Next
End Sub

dim x
Sub FlasherTimer_Timer()
    flashm 0, Flasher4Ra1
    flashm 0, Flasher4Rb1
    flashm 0, Flasher4Rb
    flash 0, Flasher4Ra

    flashm 1, Flasher5Rb1
    flashm 1, Flasher5Rb
    flash 1, Flasher5Ra

    flashm 2, Flasher6Rb1
    flashm 2, Flasher6Rb
    flash 2, Flasher6Ra

    flashm 3, Flasher7R1
    flash 3, Flasher7R

    flashm 4, Flasher2Rb
    flash 4, Flasher2Ra

    flashm 5, Flasher3R1
    flash 5, Flasher3R

    flash 6, Flasher1R
    flash 7, Flasher8R
End Sub

Sub SetFlash(nr, stat)
    FlashState(nr) = ABS(stat)
End Sub

Sub Flash(nr, object)
    Select Case FlashState(nr)
        Case 0 'off
            FlashLevel(nr) = FlashLevel(nr) - FlashSpeedDown
            If FlashLevel(nr) < 0 Then
                FlashLevel(nr) = 0
                FlashState(nr) = -1 'completely off
            End if
            'Object.alpha = FlashLevel(nr)
        Case 1 ' on
            FlashLevel(nr) = FlashLevel(nr) + FlashSpeedUp
            If FlashLevel(nr) > FlashMaxAlpha Then              '255 original JP code
                FlashLevel(nr) = FlashMaxAlpha                  '255 original JP code
                FlashState(nr) = -2 'completely on
            End if
            'Object.alpha = FlashLevel(nr)
    End Select
End Sub

Sub Flashm(nr, object) 'multiple flashers, it doesn't change the flashstate
    Select Case FlashState(nr)
        Case 0         'off
            'Object.alpha = FlashLevel(nr)
        Case 1         ' on
            'Object.alpha = FlashLevel(nr)
    End Select
End Sub

' *******************************************************************************************************
' Positional Sound Playback Functions by DJRobX, Rothbauerw and Herweh
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


'********************************************************************
'      JP's VP10 Rolling Sounds (+rothbauerw's Dropping Sounds)
'********************************************************************

Const tnob = 4 ' total number of balls
ReDim rolling(tnob)
InitRolling

Sub InitRolling
    Dim i
    For i = 0 to tnob
        rolling(i) = False
    Next
End Sub

Sub RollingTimer_Timer()
    Dim BOT, bb
    BOT = GetBalls

    ' stop the sound of deleted balls
    For bb = UBound(BOT) + 1 to tnob
        rolling(bb) = False
        StopSound("fx_ballrolling" & bb)
    Next

    ' exit the sub if no balls on the table
    If UBound(BOT) = -1 Then Exit Sub

   For bb = 0 to UBound(BOT)
        ' play the rolling sound for each ball
        If BallVel(BOT(bb) ) > 1 AND BOT(bb).z < 30 Then
            rolling(bb) = True
           PlaySound("fx_ballrolling" & bb), -1, Vol(BOT(bb)), AudioPan(BOT(bb)), 0, Pitch(BOT(bb)), 1, 0, AudioFade(BOT(bb))
        Else
           If rolling(bb) = True Then
               StopSound("fx_ballrolling" & bb)
                rolling(bb) = False
           End If
        End If

        ' play ball drop sounds
       If BOT(bb).VelZ < -1 and BOT(bb).z < 55 and BOT(bb).z > 27 Then 'height adjust for ball drop sounds
        End If
   Next
End Sub

'**********************
' Ball Collision Sound
'**********************

Sub OnBallBallCollision(ball1, ball2, velocity)
  PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 2000, AudioPan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub

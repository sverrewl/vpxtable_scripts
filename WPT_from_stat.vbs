 Option Explicit
   Randomize
 Const UseVPMModSol = 1
On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

' Thalamus 2018-07-24
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Added InitVpmFFlipsSAM
' Thalamus 2018-11-01 : Improved directional sounds
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
Const VolRH     = 1    ' Rubber hits volume.
Const VolPo     = 1    ' Rubber posts volume.
Const VolPi     = 1    ' Rubber pins volume.
Const VolTarg   = 1    ' Targets volume.
Const VolKick   = 1    ' Kicker volume.
Const VolSpin   = 1.5  ' Spinners volume.
Const VolFlip   = 1    ' Flipper volume.

'******************* Options *********************
' DMD/BAckglass Controller Setting
Const GIOnDuringAttractMode     = 0                 '1 - GI on during attract, 0 - GI off during attract
dim DivValue:DivValue           = 2                 ' Change Value to 4 if LED array does not display properly on your system

LoadVPM "01560000", "sam.VBS", 3.10

'********************
'Standard definitions
'********************

    Const cGameName = "wpt_140a" 'change the romname here

     Const UseSolenoids = 1
     Const UseLamps = 0
     Const UseSync = 1
     Const HandleMech = 0

     'Standard Sounds
     Const SSolenoidOn = "Solenoid"
     Const SSolenoidOff = ""
     Const SCoin = "coin"

'************
' Table init.
'************
   'Variables
    Dim xx
    Dim Bump1,Bump2,Bump3,Mech3bank,bsTrough,bsVUK,visibleLock,bsTEject,bsSVUK,bsRScoop
    Dim dtUDrop,dtLDropLower,dtLDropUpper,dtRDrop
    Dim PlungerIM
    Dim PMag
'
  Sub Table_Init
 vpmInit Me
'*****
    With Controller
        .GameName = cGameName
        If Err Then MsgBox "Can't start Game " & cGameName & vbNewLine & Err.Description:Exit Sub
        .SplashInfoLine = "WPT"
        .HandleKeyboard = 0
        .ShowTitle = 0
        .ShowDMDOnly = 1
        .ShowFrame = 0
        .HandleMechanics = 1
        .Hidden = 0
		.Games(cGameName).Settings.Value("sound") = 1
        On Error Resume Next
        .Run GetPlayerHWnd
        If Err Then MsgBox Err.Description
    End With

	InitVpmFFlipsSAM
    On Error Goto 0

    Const IMPowerSetting = 50
    Const IMTime = 0.6
    Set plungerIM = New cvpmImpulseP
    With plungerIM
        .InitImpulseP swplunger, IMPowerSetting, IMTime
        .Switch 23
        .Random 1.5
        .InitExitSnd SoundFX("plunger",DOFContactors), SoundFX("plunger",DOFContactors)
        .CreateEvents "plungerIM"
    End With

'**Trough
    Set bsTrough = New cvpmBallStack
    bsTrough.InitSw 0, 21, 20, 19, 18, 0, 0, 0
    bsTrough.InitKick BallRelease, 90, 8
    bsTrough.InitExitSnd SoundFX("ballrelease",DOFContactors), SoundFX("Solenoid",DOFContactors)
    bsTrough.Balls = 4

    Set bsSVUK=New cvpmBallStack
    bsSVUK.InitSw 0,3,0,0,0,0,0,0
    bsSVUK.InitKick TopLaneKicker,0,20
    bsSVUK.InitExitSnd SoundFX("warehousekick",DOFContactors), SoundFX("wireramp",DOFContactors)

    Set bsVUK=New cvpmBallStack
    bsVUK.InitSw 0,55,0,0,0,0,0,0
    bsVUK.InitKick LeftVUKTop,180,12
    bsVUK.InitExitSnd SoundFX("scoopexit",DOFContactors), SoundFX("rail",DOFContactors)

    Set bsRScoop=New cvpmBallStack
    bsRScoop.InitSw 0,49,0,0,0,0,0,0
    bsRScoop.InitKick sw49,270,32
    bsRScoop.InitExitSnd SoundFX("popperball",DOFContactors), SoundFX("rail",DOFContactors)

'**Nudging
       vpmNudge.TiltSwitch=-7
       vpmNudge.Sensitivity=1
       vpmNudge.TiltObj=Array(Bumper1b,Bumper2b,Bumper3b,LeftSlingshot,RightSlingshot)

     Set dtLDropLower = new cvpmDropTarget
     With dtLDropLower
          .Initdrop Array(sw33, sw34, sw35, sw36), Array(33, 34, 35, 36)
          .InitSnd SoundFX("droptarget",DOFDropTargets),SoundFX("resetdrop",DOFDropTargets)
      End With

     Set dtLDropUpper = new cvpmDropTarget
     With dtLDropUpper
          .Initdrop Array(sw37, sw38, sw39, sw40), Array(37, 38, 39, 40)
          .InitSnd SoundFX("droptarget",DOFDropTargets),SoundFX("resetdrop",DOFDropTargets)
      End With

     Set dtUDrop = new cvpmDropTarget
     With dtUDrop
          .Initdrop Array(sw10, sw11, sw12, sw13), Array(10, 11, 12, 13)
          .InitSnd SoundFX("droptarget",DOFDropTargets),SoundFX("resetdrop",DOFDropTargets)
      End With

     Set dtRDrop = new cvpmDropTarget
     With dtRDrop
          .Initdrop Array(sw4, sw5, sw6, sw7), Array(4, 5, 6, 7)
          .InitSnd SoundFX("droptarget",DOFDropTargets),SoundFX("resetdrop",DOFDropTargets)
      End With

      '**Main Timer init
    PinMAMETimer.Interval = PinMAMEInterval
    PinMAMETimer.Enabled = 1

    'If GIOnDuringAttractMode = 0 Then GI_AllOff
    'for each xx in GILights
    'xx.state = 0
    'Next

  End Sub

'*****Keys
Sub Table_KeyDown(ByVal Keycode)


    If Keycode = LeftFlipperKey then

    End If
    If Keycode = RightFlipperKey then

    End If
    If keycode = PlungerKey Then Plunger.Pullback:playsoundAtVol "plungerpull", plunger, 1
    If keycode = LeftTiltKey Then nudgebobble(keycode):End If
    If keycode = RightTiltKey Then nudgebobble(keycode):End If
   'If keycode = CenterTiltKey Then CenterNudge 0, 1, 25 End If
   If vpmKeyDown(keycode) Then Exit Sub
End Sub



Sub Table_KeyUp(ByVal keycode)
    If vpmKeyUp(keycode) Then Exit Sub
    If Keycode = LeftFlipperKey then

    End If
    If Keycode = RightFlipperKey then

    End If
    If Keycode = StartGameKey Then Controller.Switch(16) = 0
    If keycode = PlungerKey Then
        Plunger.Fire
        playsoundAtVol "plunger", plunger, 1
'        If(BallinPlunger = 1) then 'the ball is in the plunger lane
'            PlaySound SoundFX("Plunger2")
'        else
'            PlaySound "Plunger"
'        end if
    End If
End Sub

   'Solenoids
SolCallback(1) = "solTrough"
SolCallback(2) = "solAutofire"
SolCallback(3) = "bsSVUK.SolOut"
SolCallback(4) = "bSVUK.SolOut"
SolCallback(5) = "dtLDropLower.SolDropUp"
SolCallback(6) = "dtLDropUpper.SolDropUp"
SolCallback(7) = "dtUDrop.SolDropUp"
SolCallback(8) = "dtRDrop.SolDropUp"
''SolCallback(9) = "SolLeftPop" 'left pop bumper
''SolCallback(10) = "SolRightPop" 'right pop bumper
''SolCallback(11) = "SolBottomPop" 'top pop bumper
SolCallback(12) = "SolJailUp"
SolCallback(13) = "SolULFlipper"
SolCallback(14) = "SolURFlipper"
SolCallback(15) = "SolLFlipper"
SolCallback(16) = "SolRFlipper"
''SolCallBack(13)="vpmSolFlipper UpLeftFlipper,Nothing,"'Upper PF Left Flipper
''SolCallBack(14)="vpmSolFlipper UpRightFlipper,Nothing,"'Upper PF Right Flipper
''SolCallBack(15)="vpmSolFlipper LeftFlipper,Nothing,"'Left Flipper
''SolCallBack(16)="vpmSolFlipper RightFlipper,Flipper2,"'Right Flipper
'
''SolCallback(17) = 'left slingshot
''SolCallback(18) = 'right slingshot
SolCallback(19) = "SolJailLatch" 'jail latch
SolCallback(20) = "LRPost"
'SolCallBack(21) = "bsRScoop.SolOut"
SolCallback(21) = "ScoopOut"
SolCallback(22) = "LeftSlingFlash" 'left slingshot flasher
SolCallback(23) = "RightSlingFlash" 'right slingshot flasher
'
''SolCallback(24) = "vpmSolSound SoundFX(""knocker""),"
SolCallback(25) = "FlashLeft" 'flash left spinner
SolCallback(26) = "BackPanel1" 'back panel 1 left
SolCallback(27) = "BackPanel2" 'back panel 2
SolCallback(28) = "BackPanel3" 'back panel 3
SolCallback(29) = "BackPanel4" 'back panel 4
SolCallback(30) = "BackPanel5" 'back panel 5 right
SolCallback(31) = "FlashVUK" 'right vuk flash
SolCallback(32) = "RRDownPost" 'right ramp post
''SolCallback(sLLFlipper)="vpmSolFlipper LeftFlipper,UpLeftFlipper,"
''SolCallback(sLRFlipper)="vpmSolFlipper RightFlipper,UpRightFlipper,"

Sub SolLFlipper(Enabled)
     If Enabled Then
        PlaySoundAtVol SoundFXDOF("flipperupleft",102,DOFOn,DOFFlippers), LeftFlipper, VolFlip
        PlaySoundAtVol "flipperupleft", UpLeftFlipper, VolFlip
        LeftFlipper.RotateToEnd
        UpLeftFlipper.RotateToEnd
     Else
        PlaySoundAtVol SoundFXDOF("flipperdown",102,DOFOff,DOFFlippers), LeftFlipper, VolFlip
        PlaySoundAtVol "flipperdown", UpLeftFlipper, VolFlip
        LeftFlipper.RotateToStart
        UpLeftFlipper.RotateToStart
     End If
 End Sub

Sub SolRFlipper(Enabled)
     If Enabled Then
         PlaySoundAtVol SoundFXDOF("flipperupright",103,DOFOn,DOFFlippers), RightFlipper, VolFlip
         PlaySoundAtVol "flipperupright", UpRightFlipper, VolFlip
         RightFlipper.RotateToEnd
         UpRightFlipper.RotateToEnd
     Else
         PlaySoundAtVol SoundFXDOF("flipperdown",103,DOFOff,DOFFlippers),RightFlipper,VolFlip
         PlaySoundAtVol "flipperdown",UpRightFlipper,VolFlip
         RightFlipper.RotateToStart
         UpRightFlipper.RotateToStart
    End If
 End Sub

Sub SolULFlipper(Enabled)
     If Enabled Then
         PlaySound SoundFXDOF("",102,DOFOn,DOFFlippers)
		     PlaySoundAt "flipperupleft", UpLeftFlipper
         UpLeftFlipper.RotateToEnd
     Else
         PlaySound SoundFXDOF("",102,DOFOff,DOFFlippers)
		     PlaySoundAt "flipperdown", UpLeftFlipper
         UpLeftFlipper.RotateToStart
     End If
 End Sub

Sub SolURFlipper(Enabled)
     If Enabled Then
         PlaySound SoundFXDOF("",103,DOFOn,DOFFlippers)
		     PlaySoundAt "flipperupright", UpRightFlipper
         UpRightFlipper.RotateToEnd
     Else
         PlaySound SoundFXDOF("",103,DOFOff,DOFFlippers)
		     PlaySoundAt "flipperdown", UpRightFlipper
         UpRightFlipper.RotateToStart
    End If
 End Sub

Sub ScoopOut(enabled)
    If Enabled Then
        sw49.kick 270, 32
        Controller.Switch(49) = 0
        playsoundAtVol SoundFX("popper_ball",DOFContactors), sw49, VolKick
    End If
End Sub

Sub LeftSlingFlash(Enabled)
    If Enabled Then
        SetLamp 122, 1
        SetFlash 132, 1
        'LFLogo.image = "flipper-l2red"
    Else
        SetLamp 122, 0
        SetFlash 132, 0
        'if GI_TroughCheck < 4 then LFLogo.image = "flipper-l2" else LFLogo.image = "flipper-l2off"
    End If
End Sub

Sub RightSlingFlash(Enabled)
    If Enabled Then
        SetLamp 123, 1
        SetFlash 133, 1
        'RFLogo.image = "flipper-r2red"
    Else
        SetLamp 123, 0
        SetFlash 133, 0
        'if GI_TroughCheck < 4 then RFLogo.image = "flipper-r2" else RFLogo.image = "flipper-r2off"
    End If
End Sub

Sub FlashLeft(Enabled)
    If Enabled Then
        SetLamp 125, 1
        SetFlash 135, 1
    Else
        SetLamp 125, 0
        SetFlash 135, 0
    End If
End Sub

Sub BackPanel1(Enabled)
    If Enabled Then
        SetLamp 126, 1
        SetFlash 136, 1
        'LFLogo1.image = "flipper-l2red"
    Else
        SetLamp 126, 0
        SetFlash 136, 0
        'if GI_TroughCheck < 4 then LFLogo1.image = "flipper-l2" else LFLogo1.image = "flipper-l2off"
    End If
End Sub

Sub BackPanel2(Enabled)
    If Enabled Then
        SetLamp 127, 1
        SetFlash 137, 1
        'LFLogo1.image = "flipper-l2red"
        'RFLogo1.image = "flipper-r2red"
    Else
        SetLamp 127, 0
        SetFlash 137, 0
        'if GI_TroughCheck < 4 then LFLogo1.image = "flipper-l2" else LFLogo1.image = "flipper-l2off"
        'if GI_TroughCheck < 4 then RFLogo1.image = "flipper-r2" else RFLogo1.image = "flipper-r2off"
    End If
End Sub

Sub BackPanel3(Enabled)
    If Enabled Then
        SetLamp 128, 1
        SetFlash 138, 1
        'LFLogo1.image = "flipper-l2red"
        'RFLogo1.image = "flipper-r2red"
    Else
        SetLamp 128, 0
        SetFlash 138, 0
        'if GI_TroughCheck < 4 then LFLogo1.image = "flipper-l2" else LFLogo1.image = "flipper-l2off"
        'if GI_TroughCheck < 4 then RFLogo1.image = "flipper-r2" else RFLogo1.image = "flipper-r2off"
    End If
End Sub

Sub BackPanel4(Enabled)
    If Enabled Then
        SetLamp 129, 1
        SetFlash 139, 1
        'LFLogo1.image = "flipper-l2red"
        'RFLogo1.image = "flipper-r2red"
    Else
        SetLamp 129, 0
        SetFlash 139, 0
        'if GI_TroughCheck < 4 then LFLogo1.image = "flipper-l2" else LFLogo1.image = "flipper-l2off"
        'if GI_TroughCheck < 4 then RFLogo1.image = "flipper-r2" else RFLogo1.image = "flipper-r2off"
    End If
End Sub

Sub BackPanel5(Enabled)
    If Enabled Then
        SetLamp 130, 1
        SetFlash 140, 1
        'RFLogo1.image = "flipper-r2red"
    Else
        SetLamp 130, 0
        SetFlash 140, 0
        'if GI_TroughCheck < 4 then RFLogo1.image = "flipper-r2" else RFLogo1.image = "flipper-r2off"
    End If
End Sub

Sub FlashVUK(Enabled)
    If Enabled Then
        SetLamp 131, 1
        SetFlash 141, 1
    Else
        SetLamp 131, 0
        SetFlash 141, 0
    End If
End Sub
'
Sub SolJailLatch(Enabled)
    If Enabled Then 'close jail
        JailDiv.IsDropped = false
        JailDiv1.IsDropped = false
        JailDiv2.IsDropped = false
        Controller.Switch(63)=0
    End If
End Sub

Sub SolJailUp(Enabled)
    If Enabled Then
        JailDiv.IsDropped = true
        JailDiv1.IsDropped = true
        JailDiv2.IsDropped = true
        Controller.Switch(63)=1
    End If
End Sub

Sub RRDownPost(Enabled)
    If Enabled Then
        RightPost.IsDropped = False
    Else
        RightPost.IsDropped = True
    End If
End Sub

Sub LRPost(Enabled)
    If Enabled Then
        LeftPost.IsDropped = False
    Else
        LeftPost.IsDropped = True
    End If
End Sub

Sub solTrough(Enabled)
    If Enabled Then
        bsTrough.ExitSol_On
        vpmTimer.PulseSw 22
    End If
 End Sub

Sub solAutofire(Enabled)
    If Enabled Then
        PlungerIM.AutoFire
    End If
 End Sub


Sub LeftSlingShot_Slingshot
    'Leftsling = True
    Controller.Switch(26) = 1
    PlaySoundAtVol Soundfx("left_slingshot",DOFContactors),left1f,1:LeftSlingshot.TimerEnabled = 1
    left1f.rotatetoend
    left2f.rotatetoend
    left3f.rotatetoend
  End Sub

'Dim Leftsling:Leftsling = False

'Sub LS_Timer()
    'If Leftsling = True and Left1.ObjRotZ < -7 then Left1.ObjRotZ = Left1.ObjRotZ + 2
    'If Leftsling = False and Left1.ObjRotZ > -20 then Left1.ObjRotZ = Left1.ObjRotZ - 2
    'If Left1.ObjRotZ >= -7 then Leftsling = False
    'If Leftsling = True and Left2.ObjRotZ > -212.5 then Left2.ObjRotZ = Left2.ObjRotZ - 2
    'If Leftsling = False and Left2.ObjRotZ < -199 then Left2.ObjRotZ = Left2.ObjRotZ + 2
    'If Left2.ObjRotZ <= -212.5 then Leftsling = False
    'If Leftsling = True and Left3.TransZ > -23 then Left3.TransZ = Left3.TransZ - 4
    'If Leftsling = False and Left3.TransZ < -0 then Left3.TransZ = Left3.TransZ + 4
    'If Left3.TransZ <= -23 then Leftsling = False
    'If Leftsling = True then left1f.rotatetoend
    'If Leftsling = False then left1f.rotatetostart
    'If Leftsling = True then left2f.rotatetoend
    'If Leftsling = False then left2f.rotatetostart
    'If left1f.currentangle = -212 then Leftsling = False
'   Left1.ObjRotZ = left1f.currentangle
'   Left2.ObjRotZ = left2f.currentangle
'   Left3.TransZ = left3f.currentangle
'End Sub

 Sub LeftSlingShot_Timer:Me.TimerEnabled = 0:Controller.Switch(26) = 0:left1f.rotatetostart:left2f.rotatetostart:left3f.rotatetostart:End Sub

 Sub RightSlingShot_Slingshot
    'Rightsling = True
    Controller.Switch(27) = 1
    PlaySoundAtVol Soundfx("right_slingshot",DOFContactors),right1f,1:RightSlingshot.TimerEnabled = 1
    right1f.rotatetoend
    right2f.rotatetoend
    right3f.rotatetoend
  End Sub

 'Dim Rightsling:Rightsling = False

'Sub RS_Timer()
'   If Rightsling = True and Right1.ObjRotZ > 7 then Right1.ObjRotZ = Right1.ObjRotZ - 2
'   If Rightsling = False and Right1.ObjRotZ < 20 then Right1.ObjRotZ = Right1.ObjRotZ + 2
'   If Right1.ObjRotZ <= 7 then Rightsling = False
'   If Rightsling = True and Right2.ObjRotZ < 212.5 then Right2.ObjRotZ = Right2.ObjRotZ + 2
'   If Rightsling = False and Right2.ObjRotZ > 199 then Right2.ObjRotZ = Right2.ObjRotZ - 2
'   If Right2.ObjRotZ >= 212.5 then Rightsling = False
'   If Rightsling = True and Right3.TransZ > -23 then Right3.TransZ = Right3.TransZ - 4
'   If Rightsling = False and Right3.TransZ < -0 then Right3.TransZ = Right3.TransZ + 4
'   If Right3.TransZ <= -23 then Rightsling = False
'End Sub

 Sub RightSlingShot_Timer:Me.TimerEnabled = 0:Controller.Switch(27) = 0:right1f.rotatetostart:right2f.rotatetostart:right3f.rotatetostart:End Sub

      Sub Bumper1b_Hit
      vpmTimer.PulseSw 30
      PlaySoundAtVol SoundFX("bumperright",DOFContactors),Bumper1b,VolBump
        End Sub


      Sub Bumper2b_Hit
      vpmTimer.PulseSw 31
      PlaySoundAtVol SoundFX("bumperright",DOFContactors),Bumper2b,VolBump
       End Sub

      Sub Bumper3b_Hit
      vpmTimer.PulseSw 32
      PlaySoundAtVol SoundFX("bumperright",DOFContactors),Bumper3b,VolBump
       End Sub



Sub LaneKicker_Hit:
    bsSVUK.AddBall Me:
    playsoundAtVol "safehousehit", LaneKicker, 1
End Sub

Sub LeftVUK_Hit:
    'GI_AllOff 1000
    PlaySoundAtVol "kicker_enter_center", LeftVuk,1
    bsVUK.AddBall Me:
End Sub

Sub LeftVUKTop_Hit:
    PlaySoundAtVol "kicker_enter_center", LeftVUKTop, 1
    bsVUK.AddBall Me:
End Sub

Sub ScoopTrigger_Hit:Controller.Switch(54) = 1:playsoundAtVol SoundFX("scoopleft",DOFContactors),ScoopTrigger,1:End Sub
Sub ScoopTrigger_UnHit:Controller.Switch(54) = 0:End Sub

Sub ScoopUp_Hit
    ScoopUp.TimerEnabled = true
    'Controller.Switch(54) = true

    ScoopUp.Enabled = false
    PlaySoundAtVol SoundFX("scoopleft",DOFContactors),ScoopUp,1
    'ScoopUp.KickZ 0, 35, 1, 1
    ScoopUp.Kick 0, 30, 90
End Sub

Sub ScoopUp_Timer:Me.TimerEnabled = false:ScoopUp.Enabled = true:End Sub'Controller.Switch(54) = false:End Sub

Dim BallInJail:BallInJail=FALSE
Sub JailDiv_Hit
    If BallInJail=TRUE Then
        Controller.Switch(58)=0
        sw58.TimerEnabled=1
    Else
        vpmTimer.PulseSw 57
    End If
End Sub

Sub sw58_Hit
    BallInJail=TRUE
    Controller.Switch(58)=1
    'FlasherJail.Alpha = 255
End Sub

Sub sw58_unHit
    BallInJail=FALSE
    Controller.Switch(58)=0
    'FlasherJail.Alpha = 0
End Sub

Sub sw58_Timer
    sw58.TimerEnabled=0
    If BallInJail=TRUE Then Controller.Switch(58)=1
End Sub

Sub BallJailT_Timer()
    if BallInJail = true then
        'FlasherJail.Alpha = 255
    Else
        'FlasherJail.Alpha = 0
    End If
End Sub

Sub sw8s_Spin:vpmTimer.PulseSw 8:   PlaySoundAtVol "fx_spinner",sw8s,VolSpin:End Sub
Sub sw53s_Spin: PlaySoundAtVol "fx_spinner",sw53s,VolSpin:End Sub
Sub sw44s_Spin: PlaySoundAtVol "fx_spinner",sw44a,VolSpin:End Sub
Sub leftramps_Spin: PlaySoundAtVol "fx_spinner",leftramps,VolSpin:End Sub
Sub sw50s_Spin: PlaySoundAtVol "fx_spinner",sw50s,VolSpin:End Sub
Sub sw51s_Spin: PlaySoundAtVol "fx_spinner",sw51s,VolSpin:End Sub

Sub sw24_Hit:Me.TimerEnabled = 1:Controller.Switch(24) = 1:PlaySoundAtVol "rollover",ActiveBall, 1:End Sub
Sub sw24_Timer:Me.TimerEnabled = 0:Controller.Switch(24) = 0:End Sub
Sub sw25_Hit:Me.TimerEnabled = 1:Controller.Switch(25) = 1:PlaySoundAtVol "rollover",ActiveBall, 1:End Sub
Sub sw25_Timer:Me.TimerEnabled = 0:Controller.Switch(25) = 0:End Sub
Sub sw28_Hit:Me.TimerEnabled = 1:Controller.Switch(28) = 1:PlaySoundAtVol "rollover",ActiveBall, 1:End Sub
Sub sw28_Timer:Me.TimerEnabled = 0:Controller.Switch(28) = 0:End Sub
Sub sw29_Hit:Me.TimerEnabled = 1:Controller.Switch(29) = 1:PlaySoundAtVol "rollover",ActiveBall, 1:End Sub
Sub sw29_Timer:Me.TimerEnabled = 0:Controller.Switch(29) = 0:End Sub

Sub sw33_Hit:Me.TimerEnabled = 1:dtLDropLower.Hit 1:End Sub
Sub sw33_Timer:Me.TimerEnabled = 0:End Sub
Sub sw34_Hit:Me.TimerEnabled = 1:dtLDropLower.Hit 2:End Sub
Sub sw34_Timer:Me.TimerEnabled = 0:End Sub
Sub sw35_Hit:Me.TimerEnabled = 1:dtLDropLower.Hit 3:End Sub
Sub sw35_Timer:Me.TimerEnabled = 0:End Sub
Sub sw36_Hit:Me.TimerEnabled = 1:dtLDropLower.Hit 4:End Sub
Sub sw36_Timer:Me.TimerEnabled = 0:End Sub

Sub sw37_Hit:Me.TimerEnabled = 1:dtLDropUpper.Hit 1:End Sub
Sub sw37_Timer:Me.TimerEnabled = 0:End Sub
Sub sw38_Hit:Me.TimerEnabled = 1:dtLDropUpper.Hit 2:End Sub
Sub sw38_Timer:Me.TimerEnabled = 0:End Sub
Sub sw39_Hit:Me.TimerEnabled = 1:dtLDropUpper.Hit 3:End Sub
Sub sw39_Timer:Me.TimerEnabled = 0:End Sub
Sub sw40_Hit:Me.TimerEnabled = 1:dtLDropUpper.Hit 4:End Sub
Sub sw40_Timer:Me.TimerEnabled = 0:End Sub

Sub sw9_Hit  : Controller.Switch(9) = 1: End Sub 'right ramp enter
Sub sw9_UnHit: Controller.Switch(9) = 0: End Sub

Sub sw10_Hit:Me.TimerEnabled = 1:dtUDrop.Hit 1:End Sub
Sub sw10_Timer:Me.TimerEnabled = 0:End Sub
Sub sw11_Hit:Me.TimerEnabled = 1:dtUDrop.Hit 2:End Sub
Sub sw11_Timer:Me.TimerEnabled = 0:End Sub
Sub sw12_Hit:Me.TimerEnabled = 1:dtUDrop.Hit 3:End Sub
Sub sw12_Timer:Me.TimerEnabled = 0:End Sub
Sub sw13_Hit:Me.TimerEnabled = 1:dtUDrop.Hit 4:End Sub
Sub sw13_Timer:Me.TimerEnabled = 0:End Sub
Sub sw14_Hit:vpmTimer.PulseSw 14:End Sub

Sub sw4_Hit:Me.TimerEnabled = 1:dtRDrop.Hit 1:End Sub
Sub sw4_Timer:Me.TimerEnabled = 0:End Sub
Sub sw5_Hit:Me.TimerEnabled = 1:dtRDrop.Hit 2:End Sub
Sub sw5_Timer:Me.TimerEnabled = 0:End Sub
Sub sw6_Hit:Me.TimerEnabled = 1:dtRDrop.Hit 3:End Sub
Sub sw6_Timer:Me.TimerEnabled = 0:End Sub
Sub sw7_Hit:Me.TimerEnabled = 1:dtRDrop.Hit 4:End Sub
Sub sw7_Timer:Me.TimerEnabled = 0:End Sub

Sub sw41_Hit:vpmTimer.PulseSw 41:End Sub
Sub sw42_Hit  : vpmTimer.PulseSw 42:Me.TimerEnabled = 1:sw42p.TransX = -4: playsoundAtVol SoundFX("fx_chapa",DOFTargets),ActiveBall, VolTarg: End Sub
Sub sw42_Timer:Me.TimerEnabled = 0:sw42p.TransX = 0:End Sub
Sub sw45_Hit  : vpmTimer.PulseSw 45:Me.TimerEnabled = 1:sw45p.TransX = -4: playsoundAtVol SoundFX("fx_chapa",DOFTargets),ActiveBall, VolTarg: End Sub
Sub sw45_Timer:Me.TimerEnabled = 0:sw45p.TransX = 0:End Sub
Sub sw46_Hit  : vpmTimer.PulseSw 46:Me.TimerEnabled = 1:sw46p.TransX = -4: playsoundAtVol SoundFX("fx_chapa",DOFTargets),ActiveBall, VolTarg: End Sub
Sub sw46_Timer:Me.TimerEnabled = 0:sw46p.TransX = 0:End Sub
Sub sw47_Hit  : vpmTimer.PulseSw 47:Me.TimerEnabled = 1:sw47p.TransX = -4: playsoundAtVol SoundFX("fx_chapa",DOFTargets),ActiveBall, VolTarg: End Sub
Sub sw47_Timer:Me.TimerEnabled = 0:sw47p.TransX = 0:End Sub
Sub sw48_Hit  : vpmTimer.PulseSw 48:Me.TimerEnabled = 1:sw48p.TransX = -4: playsoundAtVol SoundFX("fx_chapa",DOFTargets),ActiveBall, VolTarg: End Sub
Sub sw48_Timer:Me.TimerEnabled = 0:sw48p.TransX = 0:End Sub
Sub sw60_Hit  : vpmTimer.PulseSw 60:Me.TimerEnabled = 1:sw60p.TransX = -4: playsoundAtVol SoundFX("fx_chapa",DOFTargets),ActiveBall, VolTarg: End Sub
Sub sw60_Timer:Me.TimerEnabled = 0:sw60p.TransX = 0:End Sub
Sub sw61_Hit  : vpmTimer.PulseSw 61:Me.TimerEnabled = 1:sw61p.TransX = -4: playsoundAtVol SoundFX("fx_chapa",DOFTargets),ActiveBall, VolTarg: End Sub
Sub sw61_Timer:Me.TimerEnabled = 0:sw61p.TransX = 0:End Sub
Sub sw62_Hit  : vpmTimer.PulseSw 62:Me.TimerEnabled = 1:sw62p.TransX = -4: playsoundAtVol SoundFX("fx_chapa",DOFTargets),ActiveBall, VolTarg: End Sub
Sub sw62_Timer:Me.TimerEnabled = 0:sw62p.TransX = 0:End Sub

'Sub sw63_Hit  : vpmTimer.PulseSw 63:Me.TimerEnabled = 1:sw63p.TransX = -4: playsound SoundFX("fx_chapa"): End Sub
'Sub sw63_Timer:Me.TimerEnabled = 0:sw63p.TransX = 0:End Sub

Sub sw43s_Spin:vpmTimer.PulseSw 43:PlaySoundAtVol "fx_spinner",sw43,VolSpin:End Sub

Sub sw44_Hit  : Controller.Switch(44) = 1: End Sub 'left orbit made
Sub sw44_UnHit: Controller.Switch(44) = 0: End Sub

'Sub sw49a_Hit:bsRScoop.AddBall Me:End Sub
Dim aBall, aZpos



Sub sw49a_Hit
    Set aBall = ActiveBall
    aZpos = 50
    sw49a.TimerInterval = 2
    sw49a.TimerEnabled = 1
    Controller.Switch(49) = 1
    playsound "kicker_enter_center" ' TODO
End Sub



Sub sw49a_Timer
    aBall.Z = aZpos
    aZpos = aZpos-2
    If aZpos <0 Then
        sw49a.TimerEnabled = 0
        sw49a.DestroyBall
        sw49.CreateBall
        sw49a.Enabled = 0
        unhittimer.Enabled = 1
    End If
End Sub

Sub unhittimer_timer
    unhittimer.enabled = 0
End Sub

Sub k3trig_hit
    sw49a.enabled = 1
end sub

Sub k3trig_unHit
    'BallInPlunger = 0
End Sub

Sub sw50_Hit  : Controller.Switch(50) = 1: End Sub 'right orbit made
Sub sw50_UnHit: Controller.Switch(50) = 0: End Sub
Sub sw51_Hit  : Controller.Switch(51) = 1: End Sub 'bssvuk exit
Sub sw51_UnHit: Controller.Switch(51) = 0: End Sub
Sub sw52_Hit  : Controller.Switch(52) = 1: End Sub 'right ramp made
Sub sw52_UnHit: Controller.Switch(52) = 0: End Sub
Sub sw53_Hit  : Controller.Switch(53) = 1: End Sub 'middle ramp
Sub sw53_UnHit: Controller.Switch(53) = 0: End Sub
Sub sw54_Hit  : Controller.Switch(54) = 1: End Sub 'left ramp opto
Sub sw54_UnHit: Controller.Switch(54) = 0: End Sub

Sub sw56_Hit  : Controller.Switch(56) = 1: End Sub 'leftvuk opto
Sub sw56_UnHit: Controller.Switch(56) = 0: End Sub

Sub sw59_Hit()
'Controller.Switch(59) = 1
Light27.state = 1
vpmTimer.PulseSw 59
vpmTimer.AddTimer 100, "BallDropSound"
of.enabled = 1
End Sub 'transfer tube

Sub of_timer()
Light27.State = 0
me.enabled = 0
End Sub

'Sub sw59_UnHit: Controller.Switch(59) = 0: Light27.state = 2: PlaySound "BallDrop": End Sub

Sub sw23_Hit  : Controller.Switch(23) = 1: PlaysoundAtVol "rollover",ActiveBall, 1: GI_TroughCheck: End Sub ' shooter lane
Sub sw23_UnHit: Controller.Switch(23) = 0: End Sub

dim MotorCallback
Set MotorCallback = GetRef("GameTimer")
Sub GameTimer
    UpdateFlipperLogos
End Sub

Sub UpdateFlipperLogo_Timer
    LFLogo.RotY = LeftFlipper.CurrentAngle
    RFlogo.RotY = RightFlipper.CurrentAngle
    LFLogo1.RotY = UpLeftFlipper.CurrentAngle
    RFlogo1.RotY = UpRightFlipper.CurrentAngle
    LeftFlipperShadow.RotZ = LeftFlipper.CurrentAngle - 122
    RightFlipperShadow.RotZ = RightFlipper.CurrentAngle -238
End Sub

Sub RLS_Timer()
    sw8p.RotX = -(sw8s.currentangle) +90
    sw43p.RotX = -(sw43s.currentangle) +90
    sw44p.RotX = -(sw44s.currentangle) +90
    sw50p.RotX = -(sw50s.currentangle) +90
    sw53p.RotX = -(sw53s.currentangle) +90
    sw51p.RotX = -(sw51s.currentangle) +90
    leftrampp.RotX = -(leftramps.currentangle) +90
    Left1.ObjRotZ = left1f.currentangle +198
    Left2.ObjRotZ = left2f.currentangle +162
    Left3.TransZ = left3f.currentangle
    Right1.ObjRotZ = Right1f.currentangle +162
    Right2.ObjRotZ = Right2f.currentangle +198
    Right3.TransZ = Right3f.currentangle
End Sub

Dim sw33up, sw34up, sw35up, sw36up, sw37up, sw38up, sw39up, sw40up, sw10up, sw11up, sw12up, sw13up, sw7up, sw6up, sw5up, sw4up
Dim Jaildown, RPDown, LPUp
Dim PrimT

Sub PrimT_Timer
    if sw33.IsDropped = True then sw33up = False else sw33up = True
    if sw34.IsDropped = True then sw34up = False else sw34up = True
    if sw35.IsDropped = True then sw35up = False else sw35up = True
    if sw36.IsDropped = True then sw36up = False else sw36up = True
    if sw37.IsDropped = True then sw37up = False else sw37up = True
    if sw38.IsDropped = True then sw38up = False else sw38up = True
    if sw39.IsDropped = True then sw39up = False else sw39up = True
    if sw40.IsDropped = True then sw40up = False else sw40up = True
    if sw10.IsDropped = True then sw10up = False else sw10up = True
    if sw11.IsDropped = True then sw11up = False else sw11up = True
    if sw12.IsDropped = True then sw12up = False else sw12up = True
    if sw13.IsDropped = True then sw13up = False else sw13up = True
    if sw7.IsDropped = True then sw7up = False else sw7up = True
    if sw6.IsDropped = True then sw6up = False else sw6up = True
    if sw5.IsDropped = True then sw5up = False else sw5up = True
    if sw4.IsDropped = True then sw4up = False else sw4up = True
    if JailDiv.IsDropped = true then Jaildown = False else Jaildown = True
    if RightPost.IsDropped = true then RPDown = False else RPDown = True
    if LeftPost.IsDropped = true then LPUp = False else LPUp = True
End Sub

Sub DT_Timer()
    If sw33up = True and sw33p.z < 25 then sw33p.z = sw33p.z + 3
    If sw34up = True and sw34p.z < 25 then sw34p.z = sw34p.z + 3
    If sw35up = True and sw35p.z < 25 then sw35p.z = sw35p.z + 3
    If sw36up = True and sw36p.z < 25 then sw36p.z = sw36p.z + 3
    If sw37up = True and sw37p.z < 25 then sw37p.z = sw37p.z + 3
    If sw38up = True and sw38p.z < 25 then sw38p.z = sw38p.z + 3
    If sw39up = True and sw39p.z < 25 then sw39p.z = sw39p.z + 3
    If sw40up = True and sw40p.z < 25 then sw40p.z = sw40p.z + 3
    If sw10up = True and sw10p.z < 25 then sw10p.z = sw10p.z + 3
    If sw11up = True and sw11p.z < 25 then sw11p.z = sw11p.z + 3
    If sw12up = True and sw12p.z < 25 then sw12p.z = sw12p.z + 3
    If sw13up = True and sw13p.z < 25 then sw13p.z = sw13p.z + 3
    If Jaildown = true and Jail1.z > 160 then Jail1.z = Jail1.z - 3
    If Jaildown = true and Jail2.z > 160 then Jail2.z = Jail2.z - 3
    If RPDown = true and RPPrim.z > 160 then RPPrim.z = RPPrim.z - 3
    If LPUp = true and LeftPostPrim.z < 0 then LeftPostPrim.z = LeftPostPrim.z + 3
    If sw7up = True and sw7p.z < 25 then sw7p.z = sw7p.z + 3
    If sw6up = True and sw6p.z < 25 then sw6p.z = sw6p.z + 3
    If sw5up = True and sw5p.z < 25 then sw5p.z = sw5p.z + 3
    If sw4up = True and sw4p.z < 25 then sw4p.z = sw4p.z + 3
    If sw33up = False and sw33p.z > -20 then sw33p.z = sw33p.z - 3
    If sw34up = False and sw34p.z > -20 then sw34p.z = sw34p.z - 3
    If sw35up = False and sw35p.z > -20 then sw35p.z = sw35p.z - 3
    If sw36up = False and sw36p.z > -20 then sw36p.z = sw36p.z - 3
    If sw37up = False and sw37p.z > -20 then sw37p.z = sw37p.z - 3
    If sw38up = False and sw38p.z > -20 then sw38p.z = sw38p.z - 3
    If sw39up = False and sw39p.z > -20 then sw39p.z = sw39p.z - 3
    If sw40up = False and sw40p.z > -20 then sw40p.z = sw40p.z - 3
    If sw10up = False and sw10p.z > -20 then sw10p.z = sw10p.z - 3
    If sw11up = False and sw11p.z > -20 then sw11p.z = sw11p.z - 3
    If sw12up = False and sw12p.z > -20 then sw12p.z = sw12p.z - 3
    If sw13up = False and sw13p.z > -20 then sw13p.z = sw13p.z - 3
    If Jaildown = False and Jail1.z < 210 then Jail1.z = Jail1.z + 3
    If Jaildown = False and Jail2.z < 210 then Jail2.z = Jail2.z + 3
    If RPDown = False and RPPrim.z < 210 then RPPrim.z = RPPrim.z + 3
    If LPUp = False and LeftPostPrim.z > -50 then LeftPostPrim.z = LeftPostPrim.z - 3
    If sw7up = False and sw7p.z > -20 then sw7p.z = sw7p.z - 3
    If sw6up = False and sw6p.z > -20 then sw6p.z = sw6p.z - 3
    If sw5up = False and sw5p.z > -20 then sw5p.z = sw5p.z - 3
    If sw4up = False and sw4p.z > -20 then sw4p.z = sw4p.z - 3
    If sw33p.z >= -20 then sw33up = False
    If sw34p.z >= -20 then sw34up = False
    If sw35p.z >= -20 then sw35up = False
    If sw36p.z >= -20 then sw36up = False
    If sw37p.z >= -20 then sw37up = False
    If sw38p.z >= -20 then sw38up = False
    If sw39p.z >= -20 then sw39up = False
    If sw40p.z >= -20 then sw40up = False
    If sw10p.z >= -20 then sw10up = False
    If sw11p.z >= -20 then sw11up = False
    If sw12p.z >= -20 then sw12up = False
    If sw13p.z >= -20 then sw13up = False
    If Jail1.z <= 210 then Jaildown = false
    If Jail2.z <= 210 then Jaildown = false
'   f RPPrim.z <= 210 then RPDown = false
    If sw7p.z >= -20 then sw7up = False
    If sw6p.z >= -20 then sw6up = False
    If sw5p.z >= -20 then sw5up = False
    If sw4p.z >= -20 then sw4up = False
End Sub


'Sub LampTimer_Timer()
'    Dim chgLamp, num, chg, ii
'    chgLamp = Controller.ChangedLamps
'    If Not IsEmpty(chgLamp) Then
'        For ii = 0 To UBound(chgLamp)
'            LampState(chgLamp(ii, 0) ) = chgLamp(ii, 1)
'            FadingLevel(chgLamp(ii, 0) ) = chgLamp(ii, 1) + 4
'           FlashState(chgLamp(ii, 0) ) = chgLamp(ii, 1)
'
'        Next
'    End If
'
'    UpdateLamps
'End Sub


'***********************************************
'***********************************************
                    ' Lamps
'***********************************************
'***********************************************



 Dim LampState(200), FadingLevel(200), FadingState(200)
Dim FlashState(200), FlashLevel(200)
Dim FlashSpeedUp, FlashSpeedDown
Dim x

AllLampsOff()
LampTimer.Interval = 40 'lamp fading speed
LampTimer.Enabled = 1
'
FlashInit()
FlasherTimer.Interval = 10 'flash fading speed
FlasherTimer.Enabled = 1

Sub LampTimer_Timer()
    Dim chgLamp, num, chg, ii
    chgLamp = Controller.ChangedLamps
    If Not IsEmpty(chgLamp) Then
        For ii = 0 To UBound(chgLamp)
            LampState(chgLamp(ii, 0) ) = chgLamp(ii, 1)
            FadingLevel(chgLamp(ii, 0) ) = chgLamp(ii, 1) + 4
            FlashState(chgLamp(ii, 0) ) = chgLamp(ii, 1)

            'GI_CheckWalkerLights chgLamp(ii, 0), chgLamp(ii, 1)
        Next
    End If

    UpdateLamps
End Sub

Sub FlashInit
    Dim i
    For i = 0 to 200
        FlashState(i) = 0
        FlashLevel(i) = 0
    Next

    FlashSpeedUp = 500   ' fast speed when turning on the flasher
    FlashSpeedDown = 100 ' slow speed when turning off the flasher, gives a smooth fading
    AllFlashOff()
End Sub

Sub AllFlashOff
    Dim i
    For i = 0 to 200
        FlashState(i) = 0
    Next
End Sub

 Sub UpdateLamps()

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
    NFadeL 32, l32
    NFadeL 33, l33
    NFadeL 34, l34
    NFadeL 35, l35
    NFadeL 36, l36
    NFadeL 37, l37
    NFadeL 38, l38
    NFadeL 39, l39
    NFadeL 40, l40
    NFadeL 41, l41
    NFadeL 42, l42
    NFadeL 43, l43
    NFadeL 44, l44
    NFadeL 45, l45
    NFadeL 46, l46
    NFadeL 47, l47
    NFadeL 48, l48
    NFadeL 49, l49
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
    NFadeL 60, l60
    NFadeL 61, l61
    NFadeL 62, l62
    NFadeL 63, l63
    NFadeL 64, l64
    NFadeL 65, l65
    NFadeL 66, l66

 NFadeL 67, L67
 NFadeL 68, L68
 NFadeL 75, L75
 NFadeL 76, L76

    NFadeL 69, l69
    NFadeL 70, l70
    NFadeL 71, l71
    NFadeL 72, l72
    NFadeL 73, l73
    NFadeL 74, l74
    NFadeL 78, l78
    NFadeL 79, l79
    NFadeL 80, l80
    NFadeLm 122, f22a
    NFadeLm 122, f22b
    NFadeL 122, f22d
    NFadeLm 123, f23a
    NFadeLm 123, f23b
    NFadeL 123, f23d
    NFadeLm 126, f26a
    NFadeLm 126, f26b
    NFadeL 126, f26d
    NFadeLm 127, f27a
    NFadeLm 127, f27b
    NFadeL 127, f27d
    NFadeLm 128, f28a
    NFadeLm 128, f28b
    NFadeL 128, f28d
    NFadeLm 129, f29a
    NFadeLm 129, f29b
    NFadeL 129, f29d
    NFadeLm 130, f30a
    NFadeLm 130, f30b
    NFadeL 130, f30d
    NFadeLm 131, f31a
    NFadeLm 131, f31b
    NFadeL 131, f31d
    NFadeLm 125, F25
    NFadeL 125, F25a
'   FlashAR 122, f122b, "rf_on", "rf_a", "rf_b", Refresh 'left slingshot flash
'   FlashAR 123, f123b, "rf_on", "rf_a", "rf_b", Refresh 'right slingshot flash
'   FlashAR 125, f125b, "rf_on", "rf_a", "rf_b", Refresh 'flash left spinner
'   FlashAR 126, f126b, "rf_on", "rf_a", "rf_b", Refresh 'back panel 1 left
'   FlashAR 127, f127b, "rf_on", "rf_a", "rf_b", Refresh 'back panel 2
'   FlashAR 128, f128b, "rf_on", "rf_a", "rf_b", Refresh 'back panel 3
'   FlashAR 129, f129b, "rf_on", "rf_a", "rf_b", Refresh 'back panel 4
'   FlashAR 130, f130b, "rf_on", "rf_a", "rf_b", Refresh 'back panel 5 right
'   FlashAR 131, f131b, "rf_on", "rf_a", "rf_b", Refresh 'flash right vuk

End Sub

Sub FadePrim(nr, pri, a, b, c, d)
    Select Case FadingLevel(nr)
        Case 2:pri.image = d:FadingLevel(nr) = 0
        Case 3:pri.image = c:FadingLevel(nr) = 1
        Case 4:pri.image = b:FadingLevel(nr) = 2
        Case 5:pri.image = a:FadingLevel(nr) = 3
    End Select
End Sub

''Lights

Sub NFadeL(nr, a)
    Select Case FadingLevel(nr)
        Case 4:a.state = 0:FadingLevel(nr) = 0
        Case 5:a.State = 1:FadingLevel(nr) = 1
    End Select
End Sub

Sub NFadeLm(nr, a)
    Select Case FadingLevel(nr)
        Case 4:a.state = 0
        Case 5:a.State = 1
    End Select
End Sub

Sub Flash(nr, object)
    Select Case FlashState(nr)
        Case 0 'off
            FlashLevel(nr) = FlashLevel(nr) - FlashSpeedDown
            If FlashLevel(nr) < 0 Then
                FlashLevel(nr) = 0
                FlashState(nr) = -1 'completely off
            End if
            Object.opacity = FlashLevel(nr)
        Case 1 ' on
            FlashLevel(nr) = FlashLevel(nr) + FlashSpeedUp
            If FlashLevel(nr) > 1000 Then
                FlashLevel(nr) = 1000
                FlashState(nr) = -2 'completely on
            End if
            Object.opacity = FlashLevel(nr)
    End Select
End Sub

Sub Flashm(nr, object) 'multiple flashers, it doesn't change the flashstate
    Select Case FlashState(nr)
        Case 0         'off
            Object.opacity = FlashLevel(nr)
        Case 1         ' on
            Object.opacity = FlashLevel(nr)
    End Select
End Sub

 Sub AllLampsOff():For x = 1 to 200:LampState(x) = 4:FadingLevel(x) = 4:Next:UpdateLamps:UpdateLamps:Updatelamps:End Sub


Sub SetLamp(nr, value)
    If value = 0 AND LampState(nr) = 0 Then Exit Sub
    If value = 1 AND LampState(nr) = 1 Then Exit Sub
    LampState(nr) = abs(value) + 4
FadingLevel(nr ) = abs(value) + 4: FadingState(nr ) = abs(value) + 4
End Sub

Sub SetFlash(nr, stat)
    FlashState(nr) = ABS(stat)
End Sub

Sub FlasherTimer_Timer()

'Flash 3, f3
'Flash 4, f4
'Flash 5, f5
'Flash 6, f6
'Flash 7, f7
'Flash 8, f8
'Flash 9, f9
'Flash 10, f10
'Flash 11, f11
'Flash 12, f12
'Flash 13, f13
'Flash 14, f14
'Flash 15, f15
'Flash 16, f16
'Flash 17, f17
'Flash 18, f18
'Flash 19, f19
'Flash 20, f20
'Flash 21, f21
'Flash 22, f22
'Flash 23, f23
'Flash 24, f24
'Flash 25, f25
'Flash 26, f26
'Flash 27, f27
'Flash 28, f28
'Flash 29, f29
'Flash 30, f30
'Flash 31, f31
'Flash 32, f32
'Flash 33, f33
'Flash 34, f34
'Flash 35, f35
'Flash 36, f36
'Flash 37, f37
'Flash 38, f38
'Flash 39, f39
'Flash 40, f40
'Flash 41, f41
'Flash 42, f42
'Flash 43, f43
'Flash 44, f44
'Flash 45, f45
'Flash 46, f46
'Flash 47, f47
'Flash 48, f48
'Flash 49, f49
'Flash 50, f50
'Flash 51, f51
'Flash 52, f52
'Flash 53, f53
'Flash 54, f54
'Flash 55, f55
'Flash 56, f56
'Flash 57, f57
'Flash 58, f58
'Flash 59, f59
'Flash 60, f60
'Flash 61, f61
'Flash 62, f62 'left pop
'Flash 63, f63
'Flash 64, f64
'Flash 65, f65
'Flash 66, f66
'Flash 67, f67
'Flash 68, f68
'Flash 69, f69
'Flash 70, f70 'right pop
'Flash 71, f71
'Flash 72, f72
'Flash 73, f73
'Flash 74, f74
'Flash 75, f75
'Flash 76, f76
'Flash 77, f77
'Flash 78, f78 'bottom pop
'Flash 79, f79
'Flash 80, f80
'
Flash 132, f22c
Flash 133, f23c
'Flash 135, f125a
'Flash 136, f126a
'Flash 137, f127a
'Flash 138, f128a
'Flash 139, f129a
'Flash 140, f130a
Flash 136, f26c
Flash 137, f27c
Flash 138, f28c
Flash 139, f29c
Flash 140, f30c
Flash 141, f31c


 End Sub


''***** GI routines

Sub Drain_Hit
    PlaySoundAtVol "balltruhe", Drain, 1
    bsTrough.AddBall Me
    Drain.TimerInterval = 200
    Drain.TimerEnabled = 1
End Sub


Sub Drain_Timer
    'Debug.print GI_TroughCheck & " " & Lampstate(4) & " " & Ballsavelight
    If GI_TroughCheck = 4 And LampState(3) = 0 Then
        'GI_AllOff 0
    End If
    If GI_TroughCheck = 3 And BallInJail = True And LampState(3) = 0 Then
        'GI_AllOff 0
    End If
    Drain.TimerEnabled = False
End Sub

Sub BallRelease_UnHit():  End Sub

Dim ballsavelight

Dim MultiballFlag

Sub UpdateGI(no, Enabled)
	Select Case no
		Case 0 'Top
			If Enabled Then
				GI_AllOn
			Else
				GI_AllOff
			End If
	End Select
End Sub

set GICallback = GetRef("UpdateGI")


Sub GI_AllOff 'Turn GI Off
    'debug.print "gioff"
    DOF 166, DOFOff
    'debug.print "GI OFF " & time
    'UpdateGI 0,0
    'RampsOff
    'FlippersOff
    dim xx
    for each xx in GILights
    xx.state = 0
    Next
    'If time > 0 Then
    '    GI_AllOnT.Interval = time
    '    GI_AllOnT.Enabled = 0
    '    GI_AllOnT.Enabled = 1
    'End If
End Sub

Sub GI_AllOn 'Turn GI On
    debug.print "gion"
    DOF 166, DOFOn
    'UpdateGI 0,8
    'RampsOn
    'FlippersOn
    dim xx
    for each xx in GILights
    xx.state = 1
    Next
End Sub

Sub GI_AllOnT_Timer 'Turn GI On timer
    'UpdateGI 0,8
    'RampsOn
    'FlippersOn
    'dim xx
    'for each xx in GILights
    'xx.state = 1
    'Next
    GI_AllOnT.Enabled = 0
End Sub

Function GI_TroughCheck
    Dim Ballcount:  Ballcount = 0
    If Controller.Switch(18) = TRUE then Ballcount = Ballcount + 1
    If Controller.Switch(19) = TRUE then Ballcount = Ballcount + 1
    If Controller.Switch(20) = TRUE then Ballcount = Ballcount + 1
    If Controller.Switch(21) = TRUE then Ballcount = Ballcount + 1
    If Ballcount < 3 Then 'Keep track of multiball mode
        MultiballFlag = 1
    Else
        MultiballFlag = 0
    End If

    GI_TroughCheck = Ballcount

    'debug.print "Troughcheck " & ballcount & " Multiball " & MultiballFlag

    If ballcount = 4 then
        GameOverTimerCheck.Enabled = 1 'no ball in play
        'debug.print timer & "Game Over?"
    Else
        GameOverTimerCheck.Enabled = 0 'ball in play
        'debug.print timer & "Game Not Over"
    End If

End Function

GameOverTimerCheck.Interval = 30000
Sub GameOverTimerCheck_Timer
    'debug.print timer & "Game Over!"
    'If GIOnDuringAttractMode = 1 Then GI_AllOn
    GameOverTimerCheck.Enabled = 0
End Sub


 Dim LED(69)

LED(0)=Array(D1,D2,D3,D4,D5,D6,D7                       )
LED(1)=Array(D8,D9,D10,D11,D12,D13,D14                  )
LED(2)=Array(D15,D16,D17,D18,D19,D20,D21                )
LED(3)=Array(D22,D23,D24,D25,D26,D27,D28                )
LED(4)=Array(D29,D30,D31,D32,D33,D34,D35                )
LED(5)=Array(D36,D37,D38,D39,D40,D41,D42                )
LED(6)=Array(D43,D44,D45,D46,D47,D48,D49                )
LED(7)=Array(D50,D51,D52,D53,D54,D55,D56                )
LED(8)=Array(D57,D58,D59,D60,D61,D62,D63                )
LED(9)=Array(D64,D65,D66,D67,D68,D69,D70                )
LED(10)=Array(D71,D72,D73,D74,D75,D76,D77               )
LED(11)=Array(D78,D79,D80,D81,D82,D83,D84               )
LED(12)=Array(D85,D86,D87,D88,D89,D90,D91               )
LED(13)=Array(D92,D93,D94,D95,D96,D97,D98               )
LED(14)=Array(D99,D100,D101,D102,D103,D104,D105         )
LED(15)=Array(D106,D107,D108,D109,D110,D111,D112        )
LED(16)=Array(D113,D114,D115,D116,D117,D118,D119        )
LED(17)=Array(D120,D121,D122,D123,D124,D125,D126        )
LED(18)=Array(D127,D128,D129,D130,D131,D132,D133        )
LED(19)=Array(D134,D135,D136,D137,D138,D139,D140        )
LED(20)=Array(D141,D142,D143,D144,D145,D146,D147        )
LED(21)=Array(D148,D149,D150,D151,D152,D153,D154        )
LED(22)=Array(D155,D156,D157,D158,D159,D160,D161        )
LED(23)=Array(D162,D163,D164,D165,D166,D167,D168        )
LED(24)=Array(D169,D170,D171,D172,D173,D174,D175        )
LED(25)=Array(D176,D177,D178,D179,D180,D181,D182        )
LED(26)=Array(D183,D184,D185,D186,D187,D188,D189        )
LED(27)=Array(D190,D191,D192,D193,D194,D195,D196        )
LED(28)=Array(D197,D198,D199,D200,D201,D202,D203        )
LED(29)=Array(D204,D205,D206,D207,D208,D209,D210        )
LED(30)=Array(D211,D212,D213,D214,D215,D216,D217        )
LED(31)=Array(D218,D219,D220,D221,D222,D223,D224        )
LED(32)=Array(D225,D226,D227,D228,D229,D230,D231        )
LED(33)=Array(D232,D233,D234,D235,D236,D237,D238        )
LED(34)=Array(D239,D240,D241,D242,D243,D244,D245        )
LED(35)=Array(D246,D247,D248,D249,D250,D251,D252        )
LED(36)=Array(D253,D254,D255,D256,D257,D258,D259        )
LED(37)=Array(D260,D261,D262,D263,D264,D265,D266        )
LED(38)=Array(D267,D268,D269,D270,D271,D272,D273        )
LED(39)=Array(D274,D275,D276,D277,D278,D279,D280        )
LED(40)=Array(D281,D282,D283,D284,D285,D286,D287        )
LED(41)=Array(D288,D289,D290,D291,D292,D293,D294        )
LED(42)=Array(D295,D296,D297,D298,D299,D300,D301        )
LED(43)=Array(D302,D303,D304,D305,D306,D307,D308        )
LED(44)=Array(D309,D310,D311,D312,D313,D314,D315        )
LED(45)=Array(D316,D317,D318,D319,D320,D321,D322        )
LED(46)=Array(D323,D324,D325,D326,D327,D328,D329        )
LED(47)=Array(D330,D331,D332,D333,D334,D335,D336        )
LED(48)=Array(D337,D338,D339,D340,D341,D342,D343        )
LED(49)=Array(D344,D345,D346,D347,D348,D349,D350        )
LED(50)=Array(D351,D352,D353,D354,D355,D356,D357        )
LED(51)=Array(D358,D359,D360,D361,D362,D363,D364        )
LED(52)=Array(D365,D366,D367,D368,D369,D370,D371        )
LED(53)=Array(D372,D373,D374,D375,D376,D377,D378        )
LED(54)=Array(D379,D380,D381,D382,D383,D384,D385        )
LED(55)=Array(D386,D387,D388,D389,D390,D391,D392        )
LED(56)=Array(D393,D394,D395,D396,D397,D398,D399        )
LED(57)=Array(D400,D401,D402,D403,D404,D405,D406        )
LED(58)=Array(D407,D408,D409,D410,D411,D412,D413        )
LED(59)=Array(D414,D415,D416,D417,D418,D419,D420        )
LED(60)=Array(D421,D422,D423,D424,D425,D426,D427        )
LED(61)=Array(D428,D429,D430,D431,D432,D433,D434        )
LED(62)=Array(D435,D436,D437,D438,D439,D440,D441        )
LED(63)=Array(D442,D443,D444,D445,D446,D447,D448        )
LED(64)=Array(D449,D450,D451,D452,D453,D454,D455        )
LED(65)=Array(D456,D457,D458,D459,D460,D461,D462        )
LED(66)=Array(D463,D464,D465,D466,D467,D468,D469        )
LED(67)=Array(D470,D471,D472,D473,D474,D475,D476        )
LED(68)=Array(D477,D478,D479,D480,D481,D482,D483        )
LED(69)=Array(D484,D485,D486,D487,D488,D489,D490        )


'Dim FlasherLED(69)
'FlasherLED(0)=Array(a1,a2,a3,a4,a5,a6,a7                       )
'FlasherLED(1)=Array(a8,a9,a10,a11,a12,a13,a14                  )
'FlasherLED(2)=Array(a15,a16,a17,a18,a19,a20,a21                )
'FlasherLED(3)=Array(a22,a23,a24,a25,a26,a27,a28                )
'FlasherLED(4)=Array(a29,a30,a31,a32,a33,a34,a35                )
'FlasherLED(5)=Array(a36,a37,a38,a39,a40,a41,a42                )
'FlasherLED(6)=Array(a43,a44,a45,a46,a47,a48,a49                )
'FlasherLED(7)=Array(a50,a51,a52,a53,a54,a55,a56                )
'FlasherLED(8)=Array(a57,a58,a59,a60,a61,a62,a63                )
'FlasherLED(9)=Array(a64,a65,a66,a67,a68,a69,a70                )
'FlasherLED(10)=Array(a71,a72,a73,a74,a75,a76,a77               )
'FlasherLED(11)=Array(a78,a79,a80,a81,a82,a83,a84               )
'FlasherLED(12)=Array(a85,a86,a87,a88,a89,a90,a91               )
'FlasherLED(13)=Array(a92,a93,a94,a95,a96,a97,a98               )
'FlasherLED(14)=Array(a99,a100,a101,a102,a103,a104,a105         )
'FlasherLED(15)=Array(a106,a107,a108,a109,a110,a111,a112        )
'FlasherLED(16)=Array(a113,a114,a115,a116,a117,a118,a119        )
'FlasherLED(17)=Array(a120,a121,a122,a123,a124,a125,a126        )
'FlasherLED(18)=Array(a127,a128,a129,a130,a131,a132,a133        )
'FlasherLED(19)=Array(a134,a135,a136,a137,a138,a139,a140        )
'FlasherLED(20)=Array(a141,a142,a143,a144,a145,a146,a147        )
'FlasherLED(21)=Array(a148,a149,a150,a151,a152,a153,a154        )
'FlasherLED(22)=Array(a155,a156,a157,a158,a159,a160,a161        )
'FlasherLED(23)=Array(a162,a163,a164,a165,a166,a167,a168        )
'FlasherLED(24)=Array(a169,a170,a171,a172,a173,a174,a175        )
'FlasherLED(25)=Array(a176,a177,a178,a179,a180,a181,a182        )
'FlasherLED(26)=Array(a183,a184,a185,a186,a187,a188,a189        )
'FlasherLED(27)=Array(a190,a191,a192,a193,a194,a195,a196        )
'FlasherLED(28)=Array(a197,a198,a199,a200,a201,a202,a203        )
'FlasherLED(29)=Array(a204,a205,a206,a207,a208,a209,a210        )
'FlasherLED(30)=Array(a211,a212,a213,a214,a215,a216,a217        )
'FlasherLED(31)=Array(a218,a219,a220,a221,a222,a223,a224        )
'FlasherLED(32)=Array(a225,a226,a227,a228,a229,a230,a231        )
'FlasherLED(33)=Array(a232,a233,a234,a235,a236,a237,a238        )
'FlasherLED(34)=Array(a239,a240,a241,a242,a243,a244,a245        )
'FlasherLED(35)=Array(a246,a247,a248,a249,a250,a251,a252        )
'FlasherLED(36)=Array(a253,a254,a255,a256,a257,a258,a259        )
'FlasherLED(37)=Array(a260,a261,a262,a263,a264,a265,a266        )
'FlasherLED(38)=Array(a267,a268,a269,a270,a271,a272,a273        )
'FlasherLED(39)=Array(a274,a275,a276,a277,a278,a279,a280        )
'FlasherLED(40)=Array(a281,a282,a283,a284,a285,a286,a287        )
'FlasherLED(41)=Array(a288,a289,a290,a291,a292,a293,a294        )
'FlasherLED(42)=Array(a295,a296,a297,a298,a299,a300,a301        )
'FlasherLED(43)=Array(a302,a303,a304,a305,a306,a307,a308        )
'FlasherLED(44)=Array(a309,a310,a311,a312,a313,a314,a315        )
'FlasherLED(45)=Array(a316,a317,a318,a319,a320,a321,a322        )
'FlasherLED(46)=Array(a323,a324,a325,a326,a327,a328,a329        )
'FlasherLED(47)=Array(a330,a331,a332,a333,a334,a335,a336        )
'FlasherLED(48)=Array(a337,a338,a339,a340,a341,a342,a343        )
'FlasherLED(49)=Array(a344,a345,a346,a347,a348,a349,a350        )
'FlasherLED(50)=Array(a351,a352,a353,a354,a355,a356,a357        )
'FlasherLED(51)=Array(a358,a359,a360,a361,a362,a363,a364        )
'FlasherLED(52)=Array(a365,a366,a367,a368,a369,a370,a371        )
'FlasherLED(53)=Array(a372,a373,a374,a375,a376,a377,a378        )
'FlasherLED(54)=Array(a379,a380,a381,a382,a383,a384,a385        )
'FlasherLED(55)=Array(a386,a387,a388,a389,a390,a391,a392        )
'FlasherLED(56)=Array(a393,a394,a395,a396,a397,a398,a399        )
'FlasherLED(57)=Array(a400,a401,a402,a403,a404,a405,a406        )
'FlasherLED(58)=Array(a407,a408,a409,a410,a411,a412,a413        )
'FlasherLED(59)=Array(a414,a415,a416,a417,a418,a419,a420        )
'FlasherLED(60)=Array(a421,a422,a423,a424,a425,a426,a427        )
'FlasherLED(61)=Array(a428,a429,a430,a431,a432,a433,a434        )
'FlasherLED(62)=Array(a435,a436,a437,a438,a439,a440,a441        )
'FlasherLED(63)=Array(a442,a443,a444,a445,a446,a447,a448        )
'FlasherLED(64)=Array(a449,a450,a451,a452,a453,a454,a455        )
'FlasherLED(65)=Array(a456,a457,a458,a459,a460,a461,a462        )
'FlasherLED(66)=Array(a463,a464,a465,a466,a467,a468,a469        )
'FlasherLED(67)=Array(a470,a471,a472,a473,a474,a475,a476        )
'FlasherLED(68)=Array(a477,a478,a479,a480,a481,a482,a483        )
'FlasherLED(69)=Array(a484,a485,a486,a487,a488,a489,a490        )


Function CheckLED (num)
    Dim tot,ii,specialcase
    if num = 13 then specialcase = 1
    For ii = (num*5 - 5) to (num*5 - 1 - specialcase)
        tot = tot + CheckLEDColumn(ii)
    Next
    'debug.print Timer & "LED " & num & " = " & tot
    CheckLED = tot
End Function

Function CheckLEDColumn (num)
    Dim tot, obj,i
    i = 1
    For each obj in LED(num)
        tot = tot + obj.State * i
        i = i * 2
    Next
    'debug.print Timer & "LEDColumn " & num & " = " & tot
    CheckLEDColumn = tot
End Function



Sub SetLED14toL
    Dim ii, src, dest, iii, obj
    For each obj in LED(65)
        obj.state = 1
    Next
    For ii = 66 to 69
        For each obj in LED(ii)
            obj.state = 1
            Exit For
        Next
    Next

'   For each obj in FlasherLED(65)
'       obj.alpha = 255
'   Next
'   For ii = 66 to 69
'       For each obj in FlasherLED(ii)
'           obj.alpha = 255
'           Exit For
'       Next
'   Next
End Sub

Sub SetLED14toT
    Dim ii, src, dest, iii, obj

    D462.state = 255
    D469.state = 255
    D470.state = 255
    D471.state = 255
    D472.state = 255
    D473.state = 255
    D474.state = 255
    D475.state = 255
    D476.state = 255
    D483.state = 255
    D490.state = 255

'   a462.alpha = 255
'   a469.alpha = 255
'   a470.alpha = 255
'   a471.alpha = 255
'   a472.alpha = 255
'   a473.alpha = 255
'   a474.alpha = 255
'   a475.alpha = 255
'   a476.alpha = 255
'   a483.alpha = 255
'   a490.alpha = 255
End Sub

Sub SetLED14toLED9
    D456.state=D281.state
    D457.state=D282.state
    D458.state=D283.state
    D459.state=D284.state
    D460.state=D285.state
    D461.state=D286.state
    D462.state=D287.state
    D463.state=D288.state
    D464.state=D289.state
    D465.state=D290.state
    D466.state=D291.state
    D467.state=D292.state
    D468.state=D293.state
    D469.state=D294.state
    D470.state=D295.state
    D471.state=D296.state
    D472.state=D297.state
    D473.state=D298.state
    D474.state=D299.state
    D475.state=D300.state
    D476.state=D301.state
    D477.state=D302.state
    D478.state=D303.state
    D479.state=D304.state
    D480.state=D305.state
    D481.state=D306.state
    D482.state=D307.state
    D483.state=D308.state
    D484.state=D309.state
    D485.state=D310.state
    D486.state=D311.state
    D487.state=D312.state
    D488.state=D313.state
    D489.state=D314.state
    D490.state=D315.state
'   a456.Alpha=a281.Alpha
'   a457.Alpha=a282.Alpha
'   a458.Alpha=a283.Alpha
'   a459.Alpha=a284.Alpha
'   a460.Alpha=a285.Alpha
'   a461.Alpha=a286.Alpha
'   a462.Alpha=a287.Alpha
'   a463.Alpha=a288.Alpha
'   a464.Alpha=a289.Alpha
'   a465.Alpha=a290.Alpha
'   a466.Alpha=a291.Alpha
'   a467.Alpha=a292.Alpha
'   a468.Alpha=a293.Alpha
'   a469.Alpha=a294.Alpha
'   a470.Alpha=a295.Alpha
'   a471.Alpha=a296.Alpha
'   a472.Alpha=a297.Alpha
'   a473.Alpha=a298.Alpha
'   a474.Alpha=a299.Alpha
'   a475.Alpha=a300.Alpha
'   a476.Alpha=a301.Alpha
'   a477.Alpha=a302.Alpha
'   a478.Alpha=a303.Alpha
'   a479.Alpha=a304.Alpha
'   a480.Alpha=a305.Alpha
'   a481.Alpha=a306.Alpha
'   a482.Alpha=a307.Alpha
'   a483.Alpha=a308.Alpha
'   a484.Alpha=a309.Alpha
'   a485.Alpha=a310.Alpha
'   a486.Alpha=a311.Alpha
'   a487.Alpha=a312.Alpha
'   a488.Alpha=a313.Alpha
'   a489.Alpha=a314.Alpha
'   a490.Alpha=a315.Alpha

End Sub
Sub ClearLED14
    Dim ii, src, dest, iii, obj
    For ii = 65 to 69
        For each obj in LED(ii)
            obj.state = 0
        Next
    Next
'   For ii = 65 to 69
'       For each obj in FlasherLED(ii)
'           obj.alpha = 0
'       Next
'   Next

End Sub

Dim Led128:Led128 = True

Sub DisplayTimer_Timer
	Dim ChgLED, ii, num, chg, stat, obj
	If Led128 Then 'If B2S not enabled, use 128 bit LED mode
		On Error Resume Next
		ChgLed = Controller.ChangedLEDs (&HFFFFFFFF, &HFFFFFFFF, &HFFFFFFFF, &HFFFFFFFF) 'displays both rows, dupe leds
		On Error Goto 0
		If Err.Number<>0 Then
			Led128 = False
			Exit Sub
		End If
		If Not IsEmpty (ChgLED) Then

            For ii = 0 To UBound (chgLED)
                'LEDs
                num = chgLED (ii, 0) : chg = chgLED (ii, 1) : stat = chgLED (ii, 2)
                For Each obj In LED (num)
                    If chg And 1 Then obj.State = stat And 1
                    chg = chg \ DivValue : stat = stat \ DivValue
                Next

                'Flashers (thanks gtxjoe!)
'               num = chgLED (ii, 0) : chg = chgLED (ii, 1) : stat = chgLED (ii, 2)
'               For Each obj In FlasherLED (num)
'                   If chg And 1 Then obj.alpha = (stat And 1)*255
'                   chg = chg \ DivValue : stat = stat \ DivValue
'               Next
            Next
        End If

    Else 'B2S enabled use 64 bit ChangedLED mode with last LED simulated

    End If
End Sub

Sub Table_exit()
    If B2SOn Then Controller.Stop
End Sub

Dim leftdrop:leftdrop = 0
Sub leftdrop1_Hit:leftdrop = 1:playsoundAtVol "wireramp",ActiveBall, 1:End Sub
Sub leftdrop2_Hit:
    If leftdrop = 1 then
        PlaySound "balldrop"
        leftdrop = 0
    End If
    StopSound "wireramp"
    'leftdrop = 0
End Sub

Dim rightdrop:rightdrop = 0
Sub rightdrop1_Hit:rightdrop = 1:playsoundAtVol "wireramp",ActiveBall, 1:End Sub
Sub rightdrop2_Hit
    If rightdrop = 1 then
        PlaySound "balldrop"
        rightdrop = 0
    End If
    StopSound "wireramp"
    'rightdrop = 0
End Sub

'''*****************************************************************************************
'''*freneticamnesic level nudge script, based on rascals nudge bobble with help from gtxjoe*
'''*     add timers and "Nudgebobble(keycode)" to left and right tilt keys to activate     *
'''*****************************************************************************************
Dim bgcharctr:bgcharctr = 2
Dim centerlocation:centerlocation = 90
Dim bgdegree:bgdegree = 7 'move +/- 7 degrees
Dim bgdurationctr:bgdurationctr = 0

Sub LevelT_Timer()
    Dim loopctr
    Level.RotAndTra7 = Level.RotAndTra7 + bgcharctr  'change rotation value by bgcharctr
    'debug.print "Degrees: " & Level.RotAndTra7 & " Max degree offset: " & bgdegree & " Cycle count: " & bgdurationctr ''debug print
    If Level.RotAndTra7 >= bgdegree + centerlocation then bgcharctr = -1:bgdurationctr = bgdurationctr + 1   'if level moves past max degrees, change direction and increate durationctr
    If Level.RotAndTra7 <= -bgdegree + centerlocation then bgcharctr = 1  'if level moves past min location, change direction
    If bgdurationctr = 4 then bgdegree = bgdegree - 2:bgdurationctr = 0 'if level has moved back and forth 5 times, decrease amount of movement by -2 and repeat by resetting durationctr
    If bgdegree <= 0 then LevelT.Enabled = False:bgdegree = 7 'if amount of movement is 0, turn off LevelT timer and reset movement back to max 7 degrees
End Sub


Sub Nudgebobble(keycode)
    If keycode = LeftTiltKey then bgcharctr = -1  'if nudge left, move in - direction
    If keycode = RightTiltKey then bgcharctr = 1  'if nudge left, move in + direction
    If keycode = CenterTiltKey then         'if nudge center, generate random number 1 or 2.  If 1 change it to -2.  use this number for initial direction
        Dim randombobble:randombobble = Int(2 * Rnd + 1)
        If randombobble = 1 then randombobble = -2
        bgcharctr = randombobble
    End If
    LevelT.Enabled = True:bgdurationctr = 0:bgdegree = 7
End Sub

Sub bobblesome_Timer()  'This looks like a free running timer that 1 out of ten times will start movement
    Dim chance
    chance = Int(10*Rnd+1)
    If chance = 5 then Nudgebobble(CenterTiltKey)
End Sub

' *********************************************************************
'                      Supporting Ball & Sound Functions
' *********************************************************************

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
    Vol = Csng(BallVel(ball) ^2 / 50)
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "table" is the name of the table
    Dim tmp
    tmp = ball.x * 2 / table.width-1
    If tmp > 0 Then
        Pan = Csng(tmp ^10)
    Else
        Pan = Csng(-((- tmp) ^10) )
    End If
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

    ' play the rolling sound for each ball
    For b = 0 to UBound(BOT)
        If BallVel(BOT(b) ) > 1 AND BOT(b).z < 30 Then
            rolling(b) = True
            PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) ), Pan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0
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
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / (VolDiv/VolCol), Pan(ball1), 0, Pitch(ball1), 0, 0
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

' The sound is played using the VOL, PAN and PITCH functions, so the volume and pitch of the sound
' will change according to the ball speed, and the PAN function will change the stereo position according
' to the position of the ball on the table.


'**************************************
' Explanation of the collision routine
'**************************************

' The collision is built in VP.
' You only need to add a Sub OnBallBallCollision(ball1, ball2, velocity) and when two balls collide they
' will call this routine. What you add in the sub is up to you. As an example is a simple Playsound with volume and paning
' depending of the speed of the collision.



Sub Pins_Hit (idx)
	PlaySound "pinhit_low", 0, Vol(ActiveBall)*VolPi, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Targets_Hit (idx)
	PlaySound "target", 0, Vol(ActiveBall)*VolTarg, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Metals_Thin_Hit (idx)
	PlaySound "metalhit_thin", 0, Vol(ActiveBall)*VolMetal, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Metals_Medium_Hit (idx)
	PlaySound "metalhit_medium", 0, Vol(ActiveBall)*VolMetal, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Metals2_Hit (idx)
	PlaySound "metalhit2", 0, Vol(ActiveBall)*VolMetal, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Gates_Hit (idx)
	PlaySound "gate4", 0, Vol(ActiveBall)*VolGates, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Spinner_Spin
	PlaySoundAtVol "fx_spinner", Spinner, VolSpin
End Sub

Sub Rubbers_Hit(idx)
    dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
    If finalspeed > 20 then
        PlaySound "fx_rubber2", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
    End if
    If finalspeed >= 6 AND finalspeed <= 20 then
        RandomSoundRubber()
    End If
End Sub

Sub Posts_Hit(idx)
    dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
    If finalspeed > 16 then
        PlaySound "fx_rubber2", 0, Vol(ActiveBall)*VolPo, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
    End if
    If finalspeed >= 6 AND finalspeed <= 16 then
        RandomSoundRubber()
    End If
End Sub

Sub RandomSoundRubber()
    Select Case Int(Rnd*3)+1
        Case 1 : PlaySound "rubber_hit_1", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
        Case 2 : PlaySound "rubber_hit_2", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
        Case 3 : PlaySound "rubber_hit_3", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
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
        Case 1 : PlaySound "flip_hit_1", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
        Case 2 : PlaySound "flip_hit_2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
        Case 3 : PlaySound "flip_hit_3", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
    End Select
End Sub

Sub trigger1_Hit()
vpmTimer.AddTimer 100, "BallDropSound"
End Sub

Sub BallDropSound(dummy):PlaySound "BallDrop":End Sub


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

Sub PlaySoundAtVol(sound, tableobj, Volume)
  PlaySound sound, 1, Volume, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

'*********************************************************************
'                     Supporting Ball & Sound Functions
'*********************************************************************

Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "Table" is the name of the table
	Dim tmp
    tmp = tableobj.y * 2 / table.height-1
    If tmp > 0 Then
		AudioFade = Csng(tmp ^10)
    Else
        AudioFade = Csng(-((- tmp) ^10) )
    End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "Table" is the name of the table
    Dim tmp
    tmp = tableobj.x * 2 / table.width-1
    If tmp > 0 Then
        AudioPan = Csng(tmp ^10)
    Else
        AudioPan = Csng(-((- tmp) ^10) )
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


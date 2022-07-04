'**********************
'  Bad Cats(1989)
' VPX table by unclewilly,Brad1X, Clark Kent, Dark
' Table Tune-Up by members of VPin Workshop, Discord
' version 2.0
'**********************
'
' Plastics: Benji & Brad1X
' Playfield & Plastics Graphics Remaster: Brad1X
' Inserts & lights: iaakki
' Physics: Benji & iaakki
' VR Stuff: Unclewilly & Sixtoe
' Debug: Sixtoe & iaakki
' Testing: Tomato & VPin Workshop discord

'*******************************
'   VPW Revisions
'*******************************
'069 - Wrd1972 - Added new flippers from Doctor Dude
'070 - Benji - Added material to flippers, exported flipper texture and toned down saturation for more 'arcade' style flippers
'071 - iaakki - Minor GI shape adjustments near flips. Livecatch tweaked, Cabinet mode added
'072 - Sixtoe - Added VR Room & built in backbox, unified timers, removed redundant stuff, trimmed lights and fixed some lighting issues, dropped the triggers, transparancy issue remains on the centre glass roulette cover for VR, needs more work on lighting
'073 - Brad1X - Updated Table Info and Script info
'074 - iaakki - MetalSides option created, POV reworked by altering offsets only, VRBlockerWall made invisible in desktop and cabinet modes
'075 - iaakki - Flashers reworked one more time..
'076 - Benji - Re-applied POV from 074 (it got changed at some point)
'077 - iaakki - Smaller haze images
'078 - Sixtoe - Added extra sideblade primitive and refactored mode switches, center glass disabled in VR.
'079 - iaakki - Some cleanup, darksides feature removed, static rendering disabled for siderails.
'080 - Skitso - Remade GI lighting, small visual tweaks.
'2.01 - Skitso - GI rework
'2.02 - iaakki - Updated to latest physics, also separated posts and rubber bands.
'2.03 - Sixtoe - Redid VR options, minimal room added as big room is "heavy", fixed broken cabinet mode, added drop holes, various other fixes.

Option Explicit
Randomize

'Cabinet mode - Will hide the rails and scale the side panels higher
Const CabinetMode = 0

'///////////////////////-----VR Room-----///////////////////////
Const VRRoom = 0  '0 - VR Room Off, 1 - Full Room, 2 - Minimal Room, 3 - Ultra Minimal

Const VolDiv = 100    ' Lower number, louder ballrolling/collition sound
Const VolCol = 10      ' Ball collition divider ( voldiv/volcol )

' The rest of the values are multipliers
'
'  .5 = lower volume
' 1.5 = higher volume

Const VolBump   = 2    ' Bumpers volume.
Const VolBall   = 0000.1    ' Ball volume.
Const VolRol    = 1    ' Rollovers volume.
Const VolFlip   = 1    ' Flipper volume.

'** These other volume adjustments are't used but couuld be added later I suppose **'
' Const VolGates  = 1    ' Gates volume.
' Const VolMetal  = 1    ' Metals volume.
' Const VolRB     = 1    ' Rubber bands volume.
' Const VolRH     = 1    ' Rubber hits volume.
' Const VolPo     = 1    ' Rubber posts volume.
' Const VolPi     = 1    ' Rubber pins volume.
' Const VolPlast  = 1    ' Plastics volume.
' Const VolTarg   = 1    ' Targets volume.
' Const VolWood   = 1    ' Woods volume.
' Const VolKick   = 1    ' Kicker volume.
' Const VolSpin   = 1.5  ' Spinners volume.


'************************************************************************

'Primitive_TigerRamp_HP.blenddisablelighting = .5
'Primitive_FishBowl_HP.blenddisablelighting = .5
'Primitive_TigerRamp_HP_DT.blenddisablelighting = .5
'Primitive_FishBowl_HP_DT.blenddisablelighting = .5
'glass.blenddisablelighting = 1
'Plastic_Edges.blenddisablelighting = .3

Dim xx, DNS
Dns = table1.NightDay

If DNS <= 5 or DNS >= 75 then
    For each xx in aGiLights:xx.intensity = xx.intensity *(1-(DNS/100)):Next
    For each xx in aAllFlashers:xx.opacity = xx.opacity *(1-(DNS/100)):Next
    For each xx in AllLamps:xx.intensity = xx.intensity *(1-(DNS/100)):Next
    For each xx in TargetDropGi:xx.intensity = xx.intensity *(1-(DNS/100)):Next
else
    If DNS <= 40 or DNS >= 80 Then
        For each xx in aGiLights:xx.intensity = xx.intensity *(.5-(DNS/100)):Next
        For each xx in aAllFlashers:xx.opacity = xx.opacity *(.5-(DNS/100)):Next
        For each xx in AllLamps:xx.intensity = xx.intensity *(.5-(DNS/100)):Next
    For each xx in TargetDropGi:xx.intensity = xx.intensity *(.5-(DNS/100)):Next
    else
        For each xx in aGiLights:xx.intensity = xx.intensity *(.7-(DNS/100)):Next
        For each xx in aAllFlashers:xx.opacity = xx.opacity *(.7-(DNS/100)):Next
        For each xx in AllLamps:xx.intensity = xx.intensity *(.7-(DNS/100)):Next
    For each xx in TargetDropGi:xx.intensity = xx.intensity *(.7-(DNS/100)):Next
    end if
end if

solGI 0

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Dim CurrentMinute ' for VR clock
Dim PlungerType
Dim VRFPCounter, VRDKCounter
VRFPCounter = 1
VRDKCounter = 1

'*********************** VR Part 1**************************************
'**********  THE CODE BELOW RUNS THE ANIMATION TIMERS FOR THE DONKEY KONG MACHINE, FIREPLACE, AND OTHER MACHINE DMD'S.  DELETE THE CODE BELOW IF YOU HAVE DELETED THOSE OBJECTS FROM THE ROOM (You can also delete the timers on the table, but you don't have To) **************

Sub TimerAnimateCard1_Timer() ' DONKEY KONG MACHINE ANIMATION
  VRDKTube.Image = "DK " & VRDKCounter
  VRDKCounter = VRDKCounter + 1
  If VRDKCounter > 21 Then
    VRDKCounter = 1
  End If
End Sub

Sub TimerAnimateCard2_Timer() ' FIREPLACE ANIMATION
  VRFire.Image = "FP " & VRFPCounter
  VRFPCounter = VRFPCounter + 1
  If VRFPCounter > 13 Then
    VRFPCounter = 1
  End If
End Sub

'working Kit Cat Clock featured in bttf.
'**********************************************************************************************

'model from the web
'thanks rob ross for exporting the pieces
'rascal clock code
'unclewilly eye and tail animation
'copy models and timers into room add code to bottom of script


Sub kktimer()
'update clock hands from Rascal vp9 clock Table
  kkhour.ObjRotY = Hour(Now()) * 30 + (Minute(Now())/2)
  kkminute.ObjRotY = (Minute(Now()) + (Second(Now())/100))*6
End Sub

dim kkCount : kkCount = 0
dim kkDir : kkDir = 1

Sub kkEyeTail_timer()
'wag tail and move eyes 1 full rotation every second
  kkCount = kkCount + kkDir
  kktail.ObjRotY = kkCount
  kkreye.ObjRotZ = -(kkCount * 2)
  kkleye.ObjRotZ = -(kkCount * 2)
  If kkCount = 15 then kkDir = -1
  If kkCount = -15 then kkDir = 1
End Sub

'***************** CODE BELOW IS FOR THE VR CLOCK.  DELETE THIS CODE IF YOU DELETE THE VR CLOCK OBJECTS *******************************

Sub ClockTimer()
  Pminutes.RotAndTra2 = (Minute(Now())+(Second(Now())/100))*6
  Phours.RotAndTra2 = Hour(Now())*30+(Minute(Now())/2)
    Pseconds.RotAndTra2 = (Second(Now()))*6
  CurrentMinute=Minute(Now())
End Sub

Dim DesktopMode
Dim UseVPMColoredDMD
Dim VarHidden, UseVPMDMD

UseVPMColoredDMD = DesktopMode
DesktopMode      = True
UseVPMDMD        = true
VarHidden        = 1
'********** End Of VR Part 1*************


LoadVPM "01550000", "S11.vbs", 3.26

Dim bsTrough, bsDog, bsTrash, dtBird, dtMilk
Const cGameName = "bcats_l5"

Const UseSolenoids = 2
Const UseLamps = 0
Const UseGI = 0
Const UseSync = 1
Const HandleMech = 0

'Standard Sounds
Const SSolenoidOn = "fx_Solenoid"
Const SSolenoidOff = ""
Const SCoin = "fx_Coin"
BSize = 25.5
BMass = 1.7

'************
' Table init.
'************
dim HiddenVar
If Table1.ShowDT = False then
    HiddenVar = 1
  Primitive_TigerRamp_HP.visible = 1
  Primitive_FishBowl_HP.visible = 1
  Primitive_TigerRamp_HP_DT.visible = 0
  Primitive_FishBowl_HP_DT.visible = 0
Else
    HiddenVar = 0
  Primitive_TigerRamp_HP.visible = 0
  Primitive_FishBowl_HP.visible = 0
  Primitive_TigerRamp_HP_DT.visible = 1
  Primitive_FishBowl_HP_DT.visible = 1
end If

Sub Table1_Init
    TableOptions()
  vpmInit me

    With Controller
        .GameName = cGameName
        If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description:Exit Sub
        .SplashInfoLine = "BadCats, Williams 1989" & vbNewLine & "VPX table by unclewilly v.1.0"
        .Games(cGameName).Settings.Value("rol") = 0 '1= rotated display, 0= normal
        .HandleMechanics = 0
        .HandleKeyboard = 0
        .ShowDMDOnly = 1
        .ShowFrame = 0
        .ShowTitle = 0
    .Hidden = HiddenVar
        If Err Then MsgBox Err.Description
    End With
    On Error Goto 0
    Controller.Run

    'Nudging
    vpmNudge.TiltSwitch = swTilt
    vpmNudge.Sensitivity = 2
    'vpmNudge.TiltObj = Array(sw60, sw61, sw62, LeftSlingshot, RightSlingShot)

    'Trough
    Set bsTrough = New cvpmBallStack
    With bsTrough
        .InitSw 0, 10, 0, 0, 0, 0, 0, 0
        .InitKick ballrelease, 90, 4
        .InitEntrySnd "fx_Solenoid", "fx_Solenoid"
        .InitExitSnd SoundFX("fx_ballrel", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
        .IsTrough = True
        .Balls = 1
    End With

    'Dog House hole
    Set bsDog = New cvpmBallStack
    With bsDog
        .InitSw 0, 22, 0, 0, 0, 0, 0, 0
        .InitKick Ralfie, 180, 25
        .InitEntrySnd "fx_kicker_enter", "fx_Solenoid"
        .InitExitSnd SoundFX("fx_kicker", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
        .IsTrough = False
    End With

    'Trash hole
    Set bsTrash = New cvpmBallStack
    With bsTrash
        .InitSaucer Bin, 24, 79, 22
        .KickForceVar = 3
        .KickAngleVar = 2
        .InitExitSnd SoundFX("fx_kicker", DOFContactors), SoundFX("fx_kicker", DOFContactors)
        .InitAddSnd SoundFX("fx_kicker_enter", DOFContactors)
        .CreateEvents "bsTrash", Bin
    End With

    'Droptargets
    set dtBird = new cvpmdroptarget
    With dtBird
        .InitDrop Array(sw25,sw26,sw27,sw28,sw29), Array(25, 26, 27, 28, 29)
        .Initsnd SoundFX("fx_droptarget", DOFContactors), SoundFX("fx_resetdrop", DOFContactors)
    End With

    set dtMilk = new cvpmdroptarget
    With dtMilk
        .InitDrop Array(sw37,sw38,sw39), Array(37, 38, 39)
        .Initsnd SoundFX("fx_droptarget", DOFContactors), SoundFX("fx_resetdrop", DOFContactors)
    End With

    ' Seafood Wheel
    Dim mSFWheelMech
    Set mSFWheelMech = New cvpmMech
    With mSFWheelMech
        .MType = vpmMechStepSol + vpmMechCircle + vpmMechLinear + vpmMechFast
        .Sol1 = 16
        .Sol2 = 15
        .Length = 200
        .Steps = 200
        .AddSw 44, 0, 99
        .Callback = GetRef("UpdateWheel")
        .Start
    End With

'Init VariTarget
    sw19w21.IsDropped = 1
    sw19w31.IsDropped = 1
    sw19w41.IsDropped = 1

    'Main Timer init
    PinMAMETimer.Interval = PinMAMEInterval
    PinMAMETimer.Enabled = 1

  center_digits()

End Sub

Sub table1_Paused:Controller.Pause = 1:End Sub
Sub table1_unPaused:Controller.Pause = 0:End Sub
Sub table1_exit:Controller.Stop:End Sub

'**********
'Timer Code
'**********

Sub FrameTimer_Timer()
  ClockTimer
  DisplayTimer
  kktimer
  VRPlungerTimer
    RollingUpdate
    BallShadowUpdate
    FlipperL.ObjRotZ=LeftFlipper.currentangle
    FlipperR.ObjRotZ=RightFlipper.currentangle
    FlipperLSh.RotZ=LeftFlipper.currentangle
    FlipperRSh.RotZ=RightFlipper.currentangle
End Sub

'**********
' Keys
'**********

Sub table1_KeyDown(ByVal Keycode)

If Keycode = LeftFlipperKey Then
   VRFlipperButtonLeft.X = VRFlipperButtonLeft.X +10
   End if

   If Keycode = RightFlipperKey Then
   VRFlipperButtonRight.X = VRFlipperButtonRight.X - 10
   End if

   If Keycode = StartGameKey Then
   StartButton.y = StartButton.y -5
   StartButton2.y = StartButton2.y -5
   End If

    If keycode = LeftTiltKey Then Nudge 90, 5:PlaySound SoundFX("fx_nudge", 0), 0, 1, -0.1, 0.25
    If keycode = RightTiltKey Then Nudge 270, 5:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0.1, 0.25
    If keycode = CenterTiltKey Then Nudge 0, 6:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0, 0.25
    If keycode = PlungerKey Then PlaySound "fx_PlungerPull", 0, 1, 0.1, 0.05:Plunger.Pullback
    If keycode = LeftFlipperKey Then
    LFPress = 1
    'Objlevel(4) = 1 : FlasherFlash4_Timer
    'Objlevel(3) = 1 : FlasherFlash3_Timer
    'Objlevel(2) = 1 : FlasherFlash2_Timer
  End If
    If keycode = RightFlipperKey Then rfpress = 1
    If vpmKeyDown(keycode) Then Exit Sub
End Sub

Sub table1_KeyUp(ByVal Keycode)
If Keycode = LeftFlipperKey Then
   VRFlipperButtonLeft.X = VRFlipperButtonLeft.X -10
   End if

   If Keycode = RightFlipperKey Then
   VRFlipperButtonRight.X = VRFlipperButtonRight.X + 10
   End if

   If Keycode = StartGameKey Then
   StartButton.y = StartButton.y +5
   StartButton2.y = StartButton2.y +5
   End If
    If keycode = PlungerKey Then PlaySound "fx_plunger", 0, 1, 0.1, 0.05:Plunger.Fire
        If keycode = LeftFlipperKey Then
        lfpress = 0
        leftflipper.eostorqueangle = EOSA
        leftflipper.eostorque = EOST
    End If
    If keycode = RightFlipperKey Then
        rfpress = 0
        rightflipper.eostorqueangle = EOSA
        rightflipper.eostorque = EOST
    End If
    If vpmKeyUp(keycode) Then Exit Sub
End Sub

Sub VRPlungerTimer
  VRPlunger.Y = 1091 + (5* Plunger.Position) -20
end sub

'*********
' Switches
'*********

'Slings & Rubbers
Dim LStep, RStep

Sub LeftSlingShot_Slingshot
    PlaySound SoundFX("fx_slingshot_left", DOFContactors), 0, 1, -0.05, 0.05
    LeftSling4.Visible = 1
    Lemk.RotX = 26
    LStep = 0
    vpmTimer.PulseSw 63
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
    PlaySound SoundFX("fx_slingshot_right", DOFContactors), 0, 1, 0.05, 0.05
    RightSling4.Visible = 1
    Remk.RotX = 26
    RStep = 0
    vpmTimer.PulseSw 64
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

'Rubbers

Sub sw40_Hit():PlaySound "fx_Rubber", 0, 1, -0.1, 0.15::vpmTimer.PulseSw 40:End Sub
Sub sw33_Hit():PlaySound "fx_Rubber", 0, 1, 0.1, 0.15::vpmTimer.PulseSw 33:End Sub
Sub sw34_Hit():PlaySound "fx_Rubber", 0, 1, 0.1, 0.15::vpmTimer.PulseSw 34:End Sub


' Bumpers
Sub sw60_Hit:vpmTimer.PulseSw 60:PlaySound SoundFX("fx_bumper1", DOFContactors), 0, 1, -0.1, 0.15:End Sub
Sub sw61_Hit:vpmTimer.PulseSw 61:PlaySound SoundFX("fx_bumper3", DOFContactors), 0, 1, 0.1, 0.15:End Sub
Sub sw62_Hit:vpmTimer.PulseSw 62:PlaySound SoundFX("fx_bumper2", DOFContactors), 0, 1, 0, 0.15:End Sub


'Rollover & Ramp Switches
Sub sw11_Hit:Controller.Switch(11) = 1:PlaySound "fx_sensor", 0, 1, -0.1, 0.15:End Sub
Sub sw11_UnHit:Controller.Switch(11) = 0:End Sub

Sub sw12_Hit:Controller.Switch(12) = 1:PlaySound "fx_sensor", 0, 1, 0, 0.15:End Sub
Sub sw12_UnHit:Controller.Switch(12) = 0:End Sub

Sub sw13_Hit:Controller.Switch(13) = 1:PlaySound "fx_sensor", 0, 1, 0, 0.15:End Sub
Sub sw13_UnHit:Controller.Switch(13) = 0:End Sub

Sub sw14_Hit:Controller.Switch(14) = 1:PlaySound "fx_sensor", 0, 1, 0, 0.15:End Sub
Sub sw14_UnHit:Controller.Switch(14) = 0:End Sub

Sub sw36_Hit:Controller.Switch(36) = 1:PlaySound "fx_sensor", 0, 1, -0.1, 0.15:End Sub
Sub sw36_UnHit:Controller.Switch(36) = 0:End Sub

Sub sw30_Hit:Controller.Switch(30) = 1:PlaySound "fx_sensor", 0, 1, -0.1, 0.15:End Sub
Sub sw30_UnHit:Controller.Switch(30) = 0:End Sub

Sub sw31_Hit:Controller.Switch(31) = 1:PlaySound "fx_sensor", 0, 1, 0.1, 0.15:End Sub
Sub sw31_UnHit:Controller.Switch(31) = 0:End Sub

Sub sw35_Hit:Controller.Switch(35) = 1:PlaySound "fx_sensor", 0, 1, -0.1, 0.15:End Sub
Sub sw35_UnHit:Controller.Switch(35) = 0:End Sub

Sub sw41_Hit:Controller.Switch(41) = 1:sw41.Timerenabled = 1:PlaySound "fx_sensor", 0, 1, 0.1, 0.15:End Sub
Sub sw41_UnHit:Controller.Switch(41) = 0:End Sub
dim sw41Dir
sw41Dir = -1
Sub sw41_Timer()
    If sw41P.ObjRotZ = 60 then sw41Dir = 5
    If sw41P.ObjRotZ = 90 then sw41Dir = -5
    sw41P.ObjRotZ = sw41P.ObjRotZ + sw41Dir
    If sw41P.ObjRotZ = 90 then sw41.timerenabled = 0
End Sub

Sub sw43_Hit:Controller.Switch(43) = 1:sw43.Timerenabled = 1:PlaySound "fx_sensor", 0, 1, 0.1, 0.15:End Sub
Sub sw43_UnHit:Controller.Switch(43) = 0:End Sub
dim sw43Dir
sw43Dir = -1
Sub sw43_Timer()
    If sw43P.ObjRotZ = 60 then sw43Dir = 5
    If sw43P.ObjRotZ = 90 then sw43Dir = -5
    sw43P.ObjRotZ = sw43P.ObjRotZ + sw43Dir
    If sw43P.ObjRotZ = 90 then sw43.timerenabled = 0
End Sub

'Ramp Gates
Sub sw16_Hit:vpmTimer.PulseSw 16:End Sub

Sub sw23_Hit:vpmTimer.PulseSw 23:End Sub

' Linear Fish Target

Sub sw19_1_Hit:vpmTimer.PulseSw 19:PlaySound SoundFX("fx_target", DOFContactors), 0, 1, -0.1, 0.15:Fisht.transY = -50:Fisht.transX = -10:sw19w11.IsDropped = 1:sw19w21.IsDropped = 0:End Sub

Sub sw19_2_Hit:vpmTimer.PulseSw 19:PlaySound SoundFX("fx_target", DOFContactors), 0, 1, -0.1, 0.15:Fisht.transY = -95:Fisht.transX = -19:sw19w21.IsDropped = 1:sw19w31.IsDropped = 0:End Sub

Sub sw19_3_Hit:vpmTimer.PulseSw 19:PlaySound SoundFX("fx_target", DOFContactors), 0, 1, -0.1, 0.15:Fisht.transY = -135:Fisht.transX = -27:sw19w31.IsDropped = 1:sw19w41.IsDropped = 0:End Sub

Sub sw19_4_Hit:vpmTimer.PulseSw 19:PlaySound SoundFX("fx_target", DOFContactors), 0, 1, -0.1, 0.15:End Sub

Sub sw19_1_UnHit
    If ActiveBall.VelY > 0 Then

        sw19w11.IsDropped = 0
        sw19w21.IsDropped = 1
    End If
End Sub

Sub sw19_2_UnHit
    If ActiveBall.VelY > 0 Then

        sw19w21.IsDropped = 0
        sw19w31.IsDropped = 1
    End If
End Sub

Sub sw19_3_UnHit
    If ActiveBall.VelY > 0 Then

        sw19w31.IsDropped = 0
        sw19w41.IsDropped = 1
    End If
End Sub

Sub sw19_4_UnHit
End Sub

Sub Fanimation_UnHit():FTTimer.Enabled = 1:End Sub

Sub FTTimer_Timer()
    If FishT.TransY < 0 Then
        FishT.TransY =FishT.TransY + 5
    FishT.TransX =FishT.TransX + 1
    Else
        FTTimer.enabled = 0
    end If
End Sub

'Droptargets VPX
Sub sw25_Hit:PlaySound SoundFX("fx_droptarget", DOFContactors), 0, 1, -0.1, 0.15:End Sub 'hit event only for the sound
Sub sw25_Dropped:dtBird.hit 1: End Sub ' Thal

Sub sw26_Hit:PlaySound SoundFX("fx_droptarget", DOFContactors), 0, 1, -0.1, 0.15:End Sub 'hit event only for the sound
Sub sw26_Dropped:dtBird.hit 2: End Sub ' Thal

Sub sw27_Hit:PlaySound SoundFX("fx_droptarget", DOFContactors), 0, 1, -0.1, 0.15:End Sub 'hit event only for the sound
Sub sw27_Dropped:dtBird.hit 3: End Sub ' Thal

Sub sw28_Hit:PlaySound SoundFX("fx_droptarget", DOFContactors), 0, 1, -0.1, 0.15:End Sub 'hit event only for the sound
Sub sw28_Dropped:dtBird.hit 4: End Sub ' Thal

Sub sw29_Hit:PlaySound SoundFX("fx_droptarget", DOFContactors), 0, 1, -0.1, 0.15:End Sub 'hit event only for the sound
Sub sw29_Dropped:dtBird.hit 5: End Sub ' Thal

Sub sw37_Hit:PlaySound SoundFX("fx_droptarget", DOFContactors), 0, 1, 0.1, 0.15:End Sub 'hit event only for the sound
Sub sw37_Dropped:dtMilk.hit 1: End Sub ' Thal

Sub sw38_Hit:PlaySound SoundFX("fx_droptarget", DOFContactors), 0, 1, 0.1, 0.15:End Sub 'hit event only for the sound
Sub sw38_Dropped:dtMilk.hit 2: End Sub ' Thal

Sub sw39_Hit:PlaySound SoundFX("fx_droptarget", DOFContactors), 0, 1, 0.1, 0.15:End Sub 'hit event only for the sound
Sub sw39_Dropped:dtMilk.hit 3: End Sub ' Thal

' Drain & holes
Sub Bin_Hit():BsTrash.AddBall 0:End Sub
Sub Drain_Hit:Playsound "fx_drain":bsTrough.AddBall Me:End Sub
Sub Ralfie_Hit:Playsound "fx_scoop", 0, 1, 0.05, 0.05:bsDog.AddBall Me:End Sub

'  Ramp Helpers
Sub LHelp_Hit():Playsound "fx_balldrop", 0, 1, -0.05, 0.05:end Sub

Sub RHelp_Hit():Playsound "fx_balldrop", 0, 1, 0.05, 0.05:end Sub
Sub WireRampSound_Hit():Playsound "WireRamp", 0, 1, 0, 0.35:end Sub
'***********
' Solenoids
'***********
' from pacdudes script
SolCallback(1) = "bsTrough.SolIn"
SolCallback(1) = "bsTrough.SolOut"
SolCallback(2) = "vpmsolsound SoundFX(""fx_knocker"",DOFKnocker),"
SolCallback(3) = "SolDogOut"
SolCallback(4) = "solMT"  'dtMilk.SolDropUp
SolCallback(5) = "bsTrash.SolOut"
SolCallback(6) = "SolBT"    'dtBird.SolDropUp
SolCallback(9) = "SolbgCat"
SolCallback(13) = "SolbgWoman"
SolCallback(10)= "SolGIBlink"
SolCallBack(23)= "SolGION"  'check to see if 10 works
SolCallBack(11)= "SolbbGION"
SolCallback(15) = "SolSFW1"
SolCallback(16) = "SolSFW"
'Flashers
SolCallback(25) = "flash125"
SolCallback(26) = "FlashYellow3"
SolCallback(27) = "flash127"
SolCallback(28) = "FlashYellow4"
SolCallback(29) = "flash129"
SolCallback(30) = "FlashRed2"
SolCallback(31) = "FlashRed2"   '"Flash131"
SolCallback(32) = "flash132"
'Solenoid Subs

Sub SolbgCat(enabled)
  If enabled Then
    'If vrCatTimer.enabled = False then vrCatTimer.enabled = true
  End If
End Sub

Sub SolbgWoman(enabled)
  If enabled Then
    If vrWomanTimer.Enabled = False then vrWomanTimer.enabled = True
  End If
End Sub

dim vrwdir:vrwdir = -1
Sub vrWomanTimer_Timer()
  If VrWoman.RotY = 0 Then vrwdir = -1
  If VrWoman.RotY = -35 Then vrwdir = 1
  VrWoman.RotY = VrWoman.RotY + vrwdir
  'VrCat.objRotY = VrCat.objRotY - 14.4
    VrCat.RotY = VrCat.RotY - 14.4
  If VrWoman.RotY = 0 then vrWomanTimer.enabled = False
End Sub

Sub vrCatTimer_Timer()
' VrCat.objRotY = VrCat.objRotY - 14.4
  'If VrCat.objRotY = -1080 then
    ''vrCatTimer.Enabled = False
    'VrCat.objRotY = 0
  'End If
End Sub

Sub SolDogOut(enabled)
    If Enabled Then
        bsDog.ExitSol_On
        SetLamp 190, 0
    End If
End Sub

Sub SolSFW(enabled)
  If enabled Then
    SetLamp 190, 1
  Else
    SetLamp 190, 0
  end If



end Sub

Sub SolSFW1(enabled)

  If enabled Then
    SetLamp 190, 1
  Else
    SetLamp 190, 0
  end If

end Sub

Sub solMT(enabled)
  If enabled Then
    dtMilk.DropSol_On
    For each xx in MTGi:xx.State = 0:next
    For each xx in MT:xx.Image = "Milk":next
  Else
  end If
end Sub

Sub solBT(enabled)
  If enabled Then
    dtBird.DropSol_On
    For each xx in BTGi:xx.State = 0:next
    For each xx in BT:xx.Image = "BirdT":next
  Else
  end If
end Sub

Sub Flash127(enabled)
  If enabled Then
    Setlamp 127, 1
  Else
    SetLamp 127, 0
  end If
end Sub

Sub Flash125(enabled)
  If enabled Then
    Setlamp 125, 1
  Else
    SetLamp 125, 0
  end If
end Sub

Sub Flash126(enabled)
  If enabled Then
    Setlamp 126, 1
  Else
    SetLamp 126, 0
  end If
end Sub

Sub Flash128(enabled)
  If enabled Then
    Setlamp 128, 1
  Else
    SetLamp 128, 0
  end If
end Sub

Sub Flash129(enabled)
  If enabled Then
    Setlamp 129, 1
  Else
    SetLamp 129, 0
  end If
end Sub

' Sub Flash130(enabled)
'   If enabled Then
'     Setlamp 130, 1
'   Else
'     SetLamp 130, 0
'   end If
' end Sub

 Sub Flash131(enabled)
   If enabled Then
     Setlamp 131, 1
   Else
     SetLamp 131, 0
   end If
 end Sub

Sub Flash132(enabled)
  If enabled Then
    Setlamp 132, 1
  Else
    SetLamp 132, 0
  end If
end Sub

Sub ACRelay(enabled)
    vpmNudge.SolGameOn enabled
End Sub

'**************
' Flipper Subs
'**************

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Const ReflipAngle = 20

Sub SolLFlipper(Enabled)
    If Enabled Then
        If leftflipper.currentangle < leftflipper.endangle + ReflipAngle Then
            PlaySoundAtVol SoundFX("TOM_Calle_ReFlip_L0" & Int(Rnd*3)+1, DOFFlippers), LeftFlipper, VolFlip
        Else
            PlaySoundAtVol SoundFX("TOM_Calle_Flipper_Attack-L01", DOFFlippers), LeftFlipper, VolFlip
            PlaySoundAtVol SoundFX("TOM_Calle_Flipper_L0" & Int(Rnd*9)+1, DOFFlippers), LeftFlipper, VolFlip
        End If
        LF.Fire
    Else
        PlaySoundAtVol SoundFX("WD_TOM_Flipper_Left_Down_" & Int(Rnd*7)+1, DOFFlippers), LeftFlipper, VolFlip
        LeftFlipper.RotateToStart
    End If
End Sub

Sub SolRFlipper(Enabled)
    If Enabled Then
        If rightflipper.currentangle < rightflipper.endangle + ReflipAngle Then
            PlaySoundAtVol SoundFX("TOM_Calle_ReFlip_R0" & Int(Rnd*3)+1, DOFFlippers), RightFlipper, VolFlip
        Else
            PlaySoundAtVol SoundFX("TOM_Calle_Flipper_Attack-R01", DOFFlippers), RightFlipper, VolFlip
            PlaySoundAtVol SoundFX("TOM_Calle_Flipper_R0" & Int(Rnd*11)+1, DOFFlippers), RightFlipper, VolFlip
        End If
        RF.Fire
    Else
        PlaySoundAtVol SoundFX("WD_TOM_Flipper_Right_Down_" & Int(Rnd*8)+1, DOFFlippers), RightFlipper, VolFlip
        RightFlipper.RotateToStart
    End If
End Sub

'Flipper Collide sounds moved to nFozzy flipper section
'Sub LeftFlipper_Collide(parm)
 '
'End Sub

'Sub RightFlipper_Collide(parm)
 '
'End Sub

'#######################################################################

''    Begin NFozzy Physics. Flipper Tricks and Rubber Dampening'
'#######################################################################'


'**************************************************
'   Flipper Collision Subs
'NOTE: COpy and overwrite collision sound from original collision subs over
'RandomSoundFlipper()' below
'**************************************************'

Sub LeftFlipper_Collide(parm)
  '****** Start Fleep*****'
  'FlipperLeftHitParm = parm/10 ' For fleep sounds'
  'If FlipperLeftHitParm > 1 Then
  ' FlipperLeftHitParm = 1
  'End If
  'FlipperLeftHitParm = FlipperUpSoundLevel * FlipperLeftHitParm
  '******* End Fleep*****'
  CheckLiveCatch Activeball, LeftFlipper, LFCount, parm
  PlaySound "fx_rubber_flipper", 0, parm / 10, -0.1, 0.15
End Sub

Sub RightFlipper_Collide(parm)
  '******Start Fleep******'
  'FlipperRightHitParm = parm/10
  'If FlipperRightHitParm > 1 Then
  ' FlipperRightHitParm = 1
  'End If
  'FlipperRightHitParm = FlipperUpSoundLevel * FlipperRightHitParm
  '******End Fleep*******'
  CheckLiveCatch Activeball, RightFlipper, RFCount, parm
  PlaySound "fx_rubber_flipper", 0, parm / 10, 0.1, 0.15
End Sub

' Don't forget to add fipper collisions to any additional flippers!

'******************************************************
'   FLIPPER CORRECTION INITIALIZATION
'******************************************************

Const LiveCatch = 16

dim LF : Set LF = New FlipperPolarity
dim RF : Set RF = New FlipperPolarity

InitPolarity

Sub InitPolarity()
        dim x, a : a = Array(LF, RF)
        for each x in a
                x.AddPoint "Ycoef", 0, RightFlipper.Y-65, 1        'disabled
                x.AddPoint "Ycoef", 1, RightFlipper.Y-11, 1
                x.enabled = True
                x.TimeDelay = 60
        Next

        AddPt "Polarity", 0, 0, 0
        AddPt "Polarity", 1, 0.05, -5
        AddPt "Polarity", 2, 0.4, -5
        AddPt "Polarity", 3, 0.6, -4.5
        AddPt "Polarity", 4, 0.65, -4.0
        AddPt "Polarity", 5, 0.7, -3.5
        AddPt "Polarity", 6, 0.75, -3.0
        AddPt "Polarity", 7, 0.8, -2.5
        AddPt "Polarity", 8, 0.85, -2.0
        AddPt "Polarity", 9, 0.9,-1.5
        AddPt "Polarity", 10, 0.95, -1.0
        AddPt "Polarity", 11, 1, -0.5
        AddPt "Polarity", 12, 1.1, 0
        AddPt "Polarity", 13, 1.3, 0

        addpt "Velocity", 0, 0,         1
        addpt "Velocity", 1, 0.16, 1.06
        addpt "Velocity", 2, 0.41,         1.05
        addpt "Velocity", 3, 0.53,         1'0.982
        addpt "Velocity", 4, 0.702, 0.968
        addpt "Velocity", 5, 0.95,  0.968
        addpt "Velocity", 6, 1.03,         0.945

        LF.Object = LeftFlipper
        LF.EndPoint = EndPointLp
        RF.Object = RightFlipper
        RF.EndPoint = EndPointRp
End Sub

Sub TriggerLF_Hit() : LF.Addball activeball : End Sub
Sub TriggerLF_UnHit() : LF.PolarityCorrect activeball : End Sub
Sub TriggerRF_Hit() : RF.Addball activeball : End Sub
Sub TriggerRF_UnHit() : RF.PolarityCorrect activeball : End Sub

'******************************************************
'     FLIPPER CORRECTION FUNCTIONS
'******************************************************

Sub AddPt(aStr, idx, aX, aY)  'debugger wrapper for adjusting flipper script in-game
  dim a : a = Array(LF, RF)
  dim x : for each x in a
    x.addpoint aStr, idx, aX, aY
  Next
End Sub

Class FlipperPolarity
  Public DebugOn, Enabled
  Private FlipAt  'Timer variable (IE 'flip at 723,530ms...)
  Public TimeDelay  'delay before trigger turns off and polarity is disabled TODO set time!
  private Flipper, FlipperStart,FlipperEnd, FlipperEndY, LR, PartialFlipCoef
  Private Balls(20), balldata(20)

  dim PolarityIn, PolarityOut
  dim VelocityIn, VelocityOut
  dim YcoefIn, YcoefOut
  Public Sub Class_Initialize
    redim PolarityIn(0) : redim PolarityOut(0) : redim VelocityIn(0) : redim VelocityOut(0) : redim YcoefIn(0) : redim YcoefOut(0)
    Enabled = True : TimeDelay = 50 : LR = 1:  dim x : for x = 0 to uBound(balls) : balls(x) = Empty : set Balldata(x) = new SpoofBall : next
  End Sub

  Public Property let Object(aInput) : Set Flipper = aInput : StartPoint = Flipper.x : End Property
  Public Property Let StartPoint(aInput) : if IsObject(aInput) then FlipperStart = aInput.x else FlipperStart = aInput : end if : End Property
  Public Property Get StartPoint : StartPoint = FlipperStart : End Property
  Public Property Let EndPoint(aInput) : FlipperEnd = aInput.x: FlipperEndY = aInput.y: End Property
  Public Property Get EndPoint : EndPoint = FlipperEnd : End Property
  Public Property Get EndPointY: EndPointY = FlipperEndY : End Property

  Public Sub AddPoint(aChooseArray, aIDX, aX, aY) 'Index #, X position, (in) y Position (out)
    Select Case aChooseArray
      case "Polarity" : ShuffleArrays PolarityIn, PolarityOut, 1 : PolarityIn(aIDX) = aX : PolarityOut(aIDX) = aY : ShuffleArrays PolarityIn, PolarityOut, 0
      Case "Velocity" : ShuffleArrays VelocityIn, VelocityOut, 1 :VelocityIn(aIDX) = aX : VelocityOut(aIDX) = aY : ShuffleArrays VelocityIn, VelocityOut, 0
      Case "Ycoef" : ShuffleArrays YcoefIn, YcoefOut, 1 :YcoefIn(aIDX) = aX : YcoefOut(aIDX) = aY : ShuffleArrays YcoefIn, YcoefOut, 0
    End Select
    if gametime > 100 then Report aChooseArray
  End Sub

  Public Sub Report(aChooseArray)   'debug, reports all coords in tbPL.text
    if not DebugOn then exit sub
    dim a1, a2 : Select Case aChooseArray
      case "Polarity" : a1 = PolarityIn : a2 = PolarityOut
      Case "Velocity" : a1 = VelocityIn : a2 = VelocityOut
      Case "Ycoef" : a1 = YcoefIn : a2 = YcoefOut
      case else :tbpl.text = "wrong string" : exit sub
    End Select
    dim str, x : for x = 0 to uBound(a1) : str = str & aChooseArray & " x: " & round(a1(x),4) & ", " & round(a2(x),4) & vbnewline : next
    tbpl.text = str
  End Sub

  Public Sub AddBall(aBall) : dim x : for x = 0 to uBound(balls) : if IsEmpty(balls(x)) then set balls(x) = aBall : exit sub :end if : Next  : End Sub

  Private Sub RemoveBall(aBall)
    dim x : for x = 0 to uBound(balls)
      if TypeName(balls(x) ) = "IBall" then
        if aBall.ID = Balls(x).ID Then
          balls(x) = Empty
          Balldata(x).Reset
        End If
      End If
    Next
  End Sub

  Public Sub Fire()
    Flipper.RotateToEnd
    processballs
  End Sub

  Public Property Get Pos 'returns % position a ball. For debug stuff.
    dim x : for x = 0 to uBound(balls)
      if not IsEmpty(balls(x) ) then
        pos = pSlope(Balls(x).x, FlipperStart, 0, FlipperEnd, 1)
      End If
    Next
  End Property

  Public Sub ProcessBalls() 'save data of balls in flipper range
    FlipAt = GameTime
    dim x : for x = 0 to uBound(balls)
      if not IsEmpty(balls(x) ) then
        balldata(x).Data = balls(x)
      End If
    Next
    PartialFlipCoef = ((Flipper.StartAngle - Flipper.CurrentAngle) / (Flipper.StartAngle - Flipper.EndAngle))
    PartialFlipCoef = abs(PartialFlipCoef-1)
  End Sub
  Private Function FlipperOn() : if gameTime < FlipAt+TimeDelay then FlipperOn = True : End If : End Function 'Timer shutoff for polaritycorrect

  Public Sub PolarityCorrect(aBall)
    if FlipperOn() then
      dim tmp, BallPos, x, IDX, Ycoef : Ycoef = 1

      'y safety Exit
      if aBall.VelY > -8 then 'ball going down
        RemoveBall aBall
        exit Sub
      end if

      'Find balldata. BallPos = % on Flipper
      for x = 0 to uBound(Balls)
        if aBall.id = BallData(x).id AND not isempty(BallData(x).id) then
          idx = x
          BallPos = PSlope(BallData(x).x, FlipperStart, 0, FlipperEnd, 1)
          if ballpos > 0.65 then  Ycoef = LinearEnvelope(BallData(x).Y, YcoefIn, YcoefOut)        'find safety coefficient 'ycoef' data
        end if
      Next

      If BallPos = 0 Then 'no ball data meaning the ball is entering and exiting pretty close to the same position, use current values.
        BallPos = PSlope(aBall.x, FlipperStart, 0, FlipperEnd, 1)
        if ballpos > 0.65 then  Ycoef = LinearEnvelope(aBall.Y, YcoefIn, YcoefOut)            'find safety coefficient 'ycoef' data
      End If

      'Velocity correction
      if not IsEmpty(VelocityIn(0) ) then
        Dim VelCoef
   :      VelCoef = LinearEnvelope(BallPos, VelocityIn, VelocityOut)

        if partialflipcoef < 1 then VelCoef = PSlope(partialflipcoef, 0, 1, 1, VelCoef)

        if Enabled then aBall.Velx = aBall.Velx*VelCoef
        if Enabled then aBall.Vely = aBall.Vely*VelCoef
      End If

      'Polarity Correction (optional now)
      if not IsEmpty(PolarityIn(0) ) then
        If StartPoint > EndPoint then LR = -1 'Reverse polarity if left flipper
        dim AddX : AddX = LinearEnvelope(BallPos, PolarityIn, PolarityOut) * LR

        if Enabled then aBall.VelX = aBall.VelX + 1 * (AddX*ycoef*PartialFlipcoef)
        'playsound "fx_knocker"
      End If
    End If
    RemoveBall aBall
  End Sub
End Class

'******************************************************
'   FLIPPER POLARITY AND RUBBER DAMPENER
'     SUPPORTING FUNCTIONS
'******************************************************

' Used for flipper correction and rubber dampeners
Sub ShuffleArray(ByRef aArray, byVal offset) 'shuffle 1d array
  dim x, aCount : aCount = 0
  redim a(uBound(aArray) )
  for x = 0 to uBound(aArray) 'Shuffle objects in a temp array
    if not IsEmpty(aArray(x) ) Then
      if IsObject(aArray(x)) then
        Set a(aCount) = aArray(x)
      Else
        a(aCount) = aArray(x)
      End If
      aCount = aCount + 1
    End If
  Next
  if offset < 0 then offset = 0
  redim aArray(aCount-1+offset) 'Resize original array
  for x = 0 to aCount-1   'set objects back into original array
    if IsObject(a(x)) then
      Set aArray(x) = a(x)
    Else
      aArray(x) = a(x)
    End If
  Next
End Sub

' Used for flipper correction and rubber dampeners
Sub ShuffleArrays(aArray1, aArray2, offset)
  ShuffleArray aArray1, offset
  ShuffleArray aArray2, offset
End Sub

' Used for flipper correction, rubber dampeners, and drop targets
Function BallSpeed(ball) 'Calculates the ball speed
    BallSpeed = SQR(ball.VelX^2 + ball.VelY^2 + ball.VelZ^2)
End Function

' Used for flipper correction and rubber dampeners
Function PSlope(Input, X1, Y1, X2, Y2)  'Set up line via two points, no clamping. Input X, output Y
  dim x, y, b, m : x = input : m = (Y2 - Y1) / (X2 - X1) : b = Y2 - m*X2
  Y = M*x+b
  PSlope = Y
End Function

' Used for flipper correction
Class spoofball
  Public X, Y, Z, VelX, VelY, VelZ, ID, Mass, Radius
  Public Property Let Data(aBall)
    With aBall
      x = .x : y = .y : z = .z : velx = .velx : vely = .vely : velz = .velz
      id = .ID : mass = .mass : radius = .radius
    end with
  End Property
  Public Sub Reset()
    x = Empty : y = Empty : z = Empty  : velx = Empty : vely = Empty : velz = Empty
    id = Empty : mass = Empty : radius = Empty
  End Sub
End Class

' Used for flipper correction and rubber dampeners
Function LinearEnvelope(xInput, xKeyFrame, yLvl)
  dim y 'Y output
  dim L 'Line
  dim ii : for ii = 1 to uBound(xKeyFrame)  'find active line
    if xInput <= xKeyFrame(ii) then L = ii : exit for : end if
  Next
  if xInput > xKeyFrame(uBound(xKeyFrame) ) then L = uBound(xKeyFrame)  'catch line overrun
  Y = pSlope(xInput, xKeyFrame(L-1), yLvl(L-1), xKeyFrame(L), yLvl(L) )

  if xInput <= xKeyFrame(lBound(xKeyFrame) ) then Y = yLvl(lBound(xKeyFrame) )  'Clamp lower
  if xInput >= xKeyFrame(uBound(xKeyFrame) ) then Y = yLvl(uBound(xKeyFrame) )  'Clamp upper

  LinearEnvelope = Y
End Function

' Used for drop targets and flipper tricks
Function Distance(ax,ay,bx,by)
  Distance = SQR((ax - bx)^2 + (ay - by)^2)
End Function

'******************************************************
'     FLIPPER TRICKS
'******************************************************

RightFlipper.timerinterval=1
Rightflipper.timerenabled=True

sub RightFlipper_timer()
  FlipperTricks LeftFlipper, LFPress, LFCount, LFEndAngle, LFState
  FlipperTricks RightFlipper, RFPress, RFCount, RFEndAngle, RFState
  FlipperNudge RightFlipper, RFEndAngle, RFEOSNudge, LeftFlipper, LFEndAngle
  FlipperNudge LeftFlipper, LFEndAngle, LFEOSNudge,  RightFlipper, RFEndAngle
end sub

Dim LFEOSNudge, RFEOSNudge

Sub FlipperNudge(Flipper1, Endangle1, EOSNudge1, Flipper2, EndAngle2)
  Dim BOT, b

  If Flipper1.currentangle = Endangle1 and EOSNudge1 <> 1 Then
    EOSNudge1 = 1
    'debug.print Flipper1.currentangle &" = "& Endangle1 &"--"& Flipper2.currentangle &" = "& EndAngle2
    If Flipper2.currentangle = EndAngle2 Then
      BOT = GetBalls
      For b = 0 to Ubound(BOT)
        If FlipperTrigger(BOT(b).x, BOT(b).y, Flipper1) Then
          'Debug.Print "ball in flip1. exit"
          exit Sub
        end If
      Next
      For b = 0 to Ubound(BOT)
        If FlipperTrigger(BOT(b).x, BOT(b).y, Flipper2) Then
          BOT(b).velx = BOT(b).velx / 1.7
          BOT(b).vely = BOT(b).vely - 1
        end If
      Next
    End If
  Else
    If Flipper1.currentangle <> EndAngle1 then
      EOSNudge1 = 0
    end if
  End If
End Sub

'*****************
' Maths
'*****************
Const PI = 3.1415927

Function dSin(degrees)
  dsin = sin(degrees * Pi/180)
End Function

Function dCos(degrees)
  dcos = cos(degrees * Pi/180)
End Function

'*************************************************
' Check ball distance from Flipper for Rem
'*************************************************

Function Distance(ax,ay,bx,by)
  Distance = SQR((ax - bx)^2 + (ay - by)^2)
End Function

Function DistancePL(px,py,ax,ay,bx,by) ' Distance between a point and a line where point is px,py
  DistancePL = ABS((by - ay)*px - (bx - ax) * py + bx*ay - by*ax)/Distance(ax,ay,bx,by)
End Function

Function Radians(Degrees)
  Radians = Degrees * PI /180
End Function

Function AnglePP(ax,ay,bx,by)
  AnglePP = Atn2((by - ay),(bx - ax))*180/PI
End Function

Function DistanceFromFlipper(ballx, bally, Flipper)
  DistanceFromFlipper = DistancePL(ballx, bally, Flipper.x, Flipper.y, Cos(Radians(Flipper.currentangle+90))+Flipper.x, Sin(Radians(Flipper.currentangle+90))+Flipper.y)
End Function

Function FlipperTrigger(ballx, bally, Flipper)
  Dim DiffAngle
  DiffAngle  = ABS(Flipper.currentangle - AnglePP(Flipper.x, Flipper.y, ballx, bally) - 90)
  If DiffAngle > 180 Then DiffAngle = DiffAngle - 360

  If DistanceFromFlipper(ballx,bally,Flipper) < 48 and DiffAngle <= 90 and Distance(ballx,bally,Flipper.x,Flipper.y) < Flipper.Length Then
    FlipperTrigger = True
  Else
    FlipperTrigger = False
  End If
End Function

' Used for drop targets and stand up targets
Function Atn2(dy, dx)
  dim pi
  pi = 4*Atn(1)

  If dx > 0 Then
    Atn2 = Atn(dy / dx)
  ElseIf dx < 0 Then
    If dy = 0 Then
      Atn2 = pi
    Else
      Atn2 = Sgn(dy) * (pi - Atn(Abs(dy / dx)))
    end if
  ElseIf dx = 0 Then
    if dy = 0 Then
      Atn2 = 0
    else
      Atn2 = Sgn(dy) * pi / 2
    end if
  End If
End Function

'*************************************************
' End - Check ball distance from Flipper for Rem
'*************************************************

dim LFPress, RFPress, LFCount, RFCount
dim LFState, RFState
dim EOST, EOSA,Frampup, FElasticity,FReturn
dim RFEndAngle, LFEndAngle

EOST = leftflipper.eostorque
EOSA = leftflipper.eostorqueangle
Frampup = LeftFlipper.rampup
FElasticity = LeftFlipper.elasticity
FReturn = LeftFlipper.return
Const EOSTnew = 0.8
Const EOSAnew = 1
Const EOSRampup = 0
Dim SOSRampup
SOSRampup = 2.5
Const LiveElasticity = 0.45
Const SOSEM = 0.815
Const EOSReturn = 0.025

LFEndAngle = Leftflipper.endangle
RFEndAngle = RightFlipper.endangle

Sub FlipperActivate(Flipper, FlipperPress)
  FlipperPress = 1
  Flipper.Elasticity = FElasticity

  Flipper.eostorque = EOST
  Flipper.eostorqueangle = EOSA
End Sub

Sub FlipperDeactivate(Flipper, FlipperPress)
  FlipperPress = 0
  Flipper.eostorqueangle = EOSA
  Flipper.eostorque = EOST*EOSReturn/FReturn


  If Abs(Flipper.currentangle) <= Abs(Flipper.endangle) + 0.1 Then
    Dim BOT, b
    BOT = GetBalls

    For b = 0 to UBound(BOT)
      If Distance(BOT(b).x, BOT(b).y, Flipper.x, Flipper.y) < 55 Then 'check for cradle
        If BOT(b).vely >= -0.4 Then BOT(b).vely = -0.4
      End If
    Next
  End If
End Sub

Sub FlipperTricks (Flipper, FlipperPress, FCount, FEndAngle, FState)
  Dim Dir
  Dir = Flipper.startangle/Abs(Flipper.startangle)  '-1 for Right Flipper

  If Abs(Flipper.currentangle) > Abs(Flipper.startangle) - 0.05 Then
    If FState <> 1 Then
      Flipper.rampup = SOSRampup
      Flipper.endangle = FEndAngle - 3*Dir
      Flipper.Elasticity = FElasticity * SOSEM
      FCount = 0
      FState = 1
    End If
  ElseIf Abs(Flipper.currentangle) <= Abs(Flipper.endangle) and FlipperPress = 1 then
    if FCount = 0 Then FCount = GameTime

    If FState <> 2 Then
      Flipper.eostorqueangle = EOSAnew
      Flipper.eostorque = EOSTnew
      Flipper.rampup = EOSRampup
      Flipper.endangle = FEndAngle
      FState = 2
    End If
  Elseif Abs(Flipper.currentangle) > Abs(Flipper.endangle) + 0.01 and FlipperPress = 1 Then
    If FState <> 3 Then
      Flipper.eostorque = EOST
      Flipper.eostorqueangle = EOSA
      Flipper.rampup = Frampup
      Flipper.Elasticity = FElasticity
      FState = 3
    End If

  End If
End Sub

Const LiveDistanceMin = 30  'minimum distance in vp units from flipper base live catch dampening will occur
Const LiveDistanceMax = 114  'maximum distance in vp units from flipper base live catch dampening will occur (tip protection)

Sub CheckLiveCatch(ball, Flipper, FCount, parm) 'Experimental new live catch
  Dim Dir
  Dir = Flipper.startangle/Abs(Flipper.startangle)    '-1 for Right Flipper
  Dim LiveCatchBounce                             'If live catch is not perfect, it won't freeze ball totally
  Dim CatchTime : CatchTime = GameTime - FCount

  if CatchTime <= LiveCatch and parm > 6 and ABS(Flipper.x - ball.x) > LiveDistanceMin and ABS(Flipper.x - ball.x) < LiveDistanceMax Then
    if CatchTime <= LiveCatch*0.5 Then            'Perfect catch only when catch time happens in the beginning of the window
      LiveCatchBounce = 0
    else
      LiveCatchBounce = Abs((LiveCatch/2) - CatchTime)  'Partial catch when catch happens a bit late
    end If

    If LiveCatchBounce = 0 and ball.velx * Dir > 0 Then ball.velx = 0
    ball.vely = LiveCatchBounce * (32 / LiveCatch) ' Multiplier for inaccuracy bounce
    ball.angmomx= 0
    ball.angmomy= 0
    ball.angmomz= 0
  End If
End Sub

'****************************************************************************
'PHYSICS DAMPENERS

'These are data mined bounce curves,
'dialed in with the in-game elasticity as much as possible to prevent angle / spin issues.
'Requires tracking ballspeed to calculate COR

Sub dPosts()
  RubbersD.dampen Activeball
End Sub

Sub dSleeves()
  SleevesD.Dampen Activeball
End Sub


dim RubbersD : Set RubbersD = new Dampener  'frubber
RubbersD.name = "Rubbers"
RubbersD.debugOn = False  'shows info in textbox "TBPout"
RubbersD.Print = False  'debug, reports in debugger (in vel, out cor)
'cor bounce curve (linear)
'for best results, try to match in-game velocity as closely as possible to the desired curve
'RubbersD.addpoint 0, 0, 0.935  'point# (keep sequential), ballspeed, CoR (elasticity)
RubbersD.addpoint 0, 0, 0.96  'point# (keep sequential), ballspeed, CoR (elasticity)
RubbersD.addpoint 1, 3.77, 0.96
RubbersD.addpoint 2, 5.76, 0.967  'dont take this as gospel. if you can data mine rubber elasticitiy, please help!
RubbersD.addpoint 3, 15.84, 0.874
RubbersD.addpoint 4, 56, 0.64 'there's clamping so interpolate up to 56 at least

dim SleevesD : Set SleevesD = new Dampener  'this is just rubber but cut down to 85%...
SleevesD.name = "Sleeves"
SleevesD.debugOn = False  'shows info in textbox "TBPout"
SleevesD.Print = False  'debug, reports in debugger (in vel, out cor)
SleevesD.CopyCoef RubbersD, 0.85

Class Dampener
  Public Print, debugOn 'tbpOut.text
  public name, Threshold  'Minimum threshold. Useful for Flippers, which don't have a hit threshold.
  Public ModIn, ModOut
  Private Sub Class_Initialize : redim ModIn(0) : redim Modout(0): End Sub

  Public Sub AddPoint(aIdx, aX, aY)
    ShuffleArrays ModIn, ModOut, 1 : ModIn(aIDX) = aX : ModOut(aIDX) = aY : ShuffleArrays ModIn, ModOut, 0
    if gametime > 100 then Report
  End Sub

  public sub Dampen(aBall)
    if threshold then if BallSpeed(aBall) < threshold then exit sub end if end if
    dim RealCOR, DesiredCOR, str, coef
    DesiredCor = LinearEnvelope(cor.ballvel(aBall.id), ModIn, ModOut )
    RealCOR = BallSpeed(aBall) / cor.ballvel(aBall.id)
    coef = desiredcor / realcor
    if debugOn then str = name & " in vel:" & round(cor.ballvel(aBall.id),2 ) & vbnewline & "desired cor: " & round(desiredcor,4) & vbnewline & _
    "actual cor: " & round(realCOR,4) & vbnewline & "ballspeed coef: " & round(coef, 3) & vbnewline
    if Print then debug.print Round(cor.ballvel(aBall.id),2) & ", " & round(desiredcor,3)

    aBall.velx = aBall.velx * coef : aBall.vely = aBall.vely * coef
    'playsound "fx_knocker"
    if debugOn then TBPout.text = str
  End Sub

  Public Sub CopyCoef(aObj, aCoef) 'alternative addpoints, copy with coef
    dim x : for x = 0 to uBound(aObj.ModIn)
      addpoint x, aObj.ModIn(x), aObj.ModOut(x)*aCoef
    Next
  End Sub


  Public Sub Report()   'debug, reports all coords in tbPL.text
    if not debugOn then exit sub
    dim a1, a2 : a1 = ModIn : a2 = ModOut
    dim str, x : for x = 0 to uBound(a1) : str = str & x & ": " & round(a1(x),4) & ", " & round(a2(x),4) & vbnewline : next
    TBPout.text = str
  End Sub

End Class

'******************************************************
'   TRACK ALL BALL VELOCITIES
'     FOR RUBBER DAMPENER AND DROP TARGETS
'******************************************************

dim cor : set cor = New CoRTracker

Class CoRTracker
  public ballvel, ballvelx, ballvely

  Private Sub Class_Initialize : redim ballvel(0) : redim ballvelx(0): redim ballvely(0) : End Sub

  Public Sub Update() 'tracks in-ball-velocity
    dim str, b, AllBalls, highestID : allBalls = getballs

    for each b in allballs
      if b.id >= HighestID then highestID = b.id
    Next

    if uBound(ballvel) < highestID then redim ballvel(highestID)  'set bounds
    if uBound(ballvelx) < highestID then redim ballvelx(highestID)  'set bounds
    if uBound(ballvely) < highestID then redim ballvely(highestID)  'set bounds

    for each b in allballs
      ballvel(b.id) = BallSpeed(b)
      ballvelx(b.id) = b.velx
      ballvely(b.id) = b.vely
    Next
  End Sub
End Class

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

Sub RDampen_Timer()
Cor.Update
End Sub

'*****************************************************************************************************
'*******************************************************************************************************
'END nFOZZY Physics'


'SeaFoodWheel Based on jp's script based on cyclone script
Dim SFWSpin
SFWSpin = 0
Sub UpdateWheel(aNewPos, aSpeed, aLastPos)
    ' 360/200= 1.8
    If aNewPos <> aLastPos then
        SFWheel.ObjRotZ = aNewPos * 1.8
    End If
End Sub

 dim FlippersEnabled


' ***************************************************************************
'       BASIC FSS(SS TYPE0) 2x16 solid state character display SETUP CODE
' ****************************************************************************
Sub center_digits()

Dim xoff, yoff, zoff, xcen, ycen, ii, xx, yy, xfact, yfact, obj, xrot, zscale
xoff = 528 ' xoffset of destination (screen coords)
yoff = -5.5 ' yoffset of destination (screen coords)
zoff = 822 ' zoffset of destination (screen coords)
xrot = -86
zscale = 0.19

xcen =(1133 /2) - (53 / 2)
ycen = (1183 /2 ) + (133 /2)
yfact =80 'y fudge factor (ycen was wrong so fix)
xfact =80


for ii =0 to 31
  For Each obj In Digits(ii)
  xx =obj.x

  obj.x = (xoff -xcen) + (xx * 0.82) +xfact
  yy = obj.y ' get the yoffset before it is changed
  obj.y =yoff

    If(yy < 0.) then
    yy = yy * -1
    end if

  obj.height =( zoff - ycen) + yy - (yy * zscale) + yfact

  obj.rotx = xrot
  Next
  Next
end sub

Dim Digits(32)
Digits(0)=Array(ax00, ax05, ax0c, ax0d, ax08, ax01, ax06, ax0e, ax02, ax03, ax04, ax07, ax0b, ax0a, ax09, ax0f)
 Digits(1)=Array(ax10, ax15, ax1c, ax1d, ax18, ax11, ax16, ax1e, ax12, ax13, ax14, ax17, ax1b, ax1a, ax19, ax1f)
 Digits(2)=Array(ax20, ax25, ax2c, ax2d, ax28, ax21, ax26, ax2e, ax22, ax23, ax24, ax27, ax2b, ax2a, ax29, ax2f)
 Digits(3)=Array(ax30, ax35, ax3c, ax3d, ax38, ax31, ax36, ax3e, ax32, ax33, ax34, ax37, ax3b, ax3a, ax39, ax3f)
 Digits(4)=Array(ax40, ax45, ax4c, ax4d, ax48, ax41, ax46, ax4e, ax42, ax43, ax44, ax47, ax4b, ax4a, ax49, ax4f)
 Digits(5)=Array(ax50, ax55, ax5c, ax5d, ax58, ax51, ax56, ax5e, ax52, ax53, ax54, ax57, ax5b, ax5a, ax59, ax5f)
 Digits(6)=Array(ax60, ax65, ax6c, ax6d, ax68, ax61, ax66, ax6e, ax62, ax63, ax64, ax67, ax6b, ax6a, ax69, ax6f)
 Digits(7)=Array(ax70, ax75, ax7c, ax7d, ax78, ax71, ax76, ax7e, ax72, ax73, ax74, ax77, ax7b, ax7a, ax79, ax7f)
 Digits(8)=Array(ax80, ax85, ax8c, ax8d, ax88, ax81, ax86, ax8e, ax82, ax83, ax84, ax87, ax8b, ax8a, ax89, ax8f)
 Digits(9)=Array(ax90, ax95, ax9c, ax9d, ax98, ax91, ax96, ax9e, ax92, ax93, ax94, ax97, ax9b, ax9a, ax99, ax9f)
 Digits(10)=Array(axa0, axa5, axac, axad, axa8, axa1, axa6, axae, axa2, axa3, axa4, axa7, axab, axaa, axa9, axaf)
 Digits(11)=Array(axb0, axb5, axbc, axbd, axb8, axb1, axb6, axbe, axb2, axb3, axb4, axb7, axbb, axba, axb9, axbf)
 Digits(12)=Array(axc0, axc5, axcc, axcd, axc8, axc1, axc6, axce, axc2, axc3, axc4, axc7, axcb, axca, axc9, axcf)
 Digits(13)=Array(axd0, axd5, axdc, axdd, axd8, axd1, axd6, axde, axd2, axd3, axd4, axd7, axdb, axda, axd9, axdf)
 Digits(14)=Array(axe0, axe5, axec, axed, axe8, axe1, axe6, axee, axe2, axe3, axe4, axe7, axeb, axea, axe9, axef)
 Digits(15)=Array(axf0, axf5, axfc, axfd, axf8, axf1, axf6, axfe, axf2, axf3, axf4, axf7, axfb, axfa, axf9, axff)

 Digits(16)=Array(bx00, bx05, bx0c, bx0d, bx08, bx01, bx06, bx0e, bx02, bx03, bx04, bx07, bx0b, bx0a, bx09, bx0f)
 Digits(17)=Array(bx10, bx15, bx1c, bx1d, bx18, bx11, bx16, bx1e, bx12, bx13, bx14, bx17, bx1b, bx1a, bx19, bx1f)
 Digits(18)=Array(bx20, bx25, bx2c, bx2d, bx28, bx21, bx26, bx2e, bx22, bx23, bx24, bx27, bx2b, bx2a, bx29, bx2f)
 Digits(19)=Array(bx30, bx35, bx3c, bx3d, bx38, bx31, bx36, bx3e, bx32, bx33, bx34, bx37, bx3b, bx3a, bx39, bx3f)
 Digits(20)=Array(bx40, bx45, bx4c, bx4d, bx48, bx41, bx46, bx4e, bx42, bx43, bx44, bx47, bx4b, bx4a, bx49, bx4f)
 Digits(21)=Array(bx50, bx55, bx5c, bx5d, bx58, bx51, bx56, bx5e, bx52, bx53, bx54, bx57, bx5b, bx5a, bx59, bx5f)
 Digits(22)=Array(bx60, bx65, bx6c, bx6d, bx68, bx61, bx66, bx6e, bx62, bx63, bx64, bx67, bx6b, bx6a, bx69, bx6f)
 Digits(23)=Array(bx70, bx75, bx7c, bx7d, bx78, bx71, bx76, bx7e, bx72, bx73, bx74, bx77, bx7b, bx7a, bx79, bx7f)
 Digits(24)=Array(bx80, bx85, bx8c, bx8d, bx88, bx81, bx86, bx8e, bx82, bx83, bx84, bx87, bx8b, bx8a, bx89, bx8f)
 Digits(25)=Array(bx90, bx95, bx9c, bx9d, bx98, bx91, bx96, bx9e, bx92, bx93, bx94, bx97, bx9b, bx9a, bx99, bx9f)
 Digits(26)=Array(bxa0, bxa5, bxac, bxad, bxa8, bxa1, bxa6, bxae, bxa2, bxa3, bxa4, bxa7, bxab, bxaa, bxa9, bxaf)
 Digits(27)=Array(bxb0, bxb5, bxbc, bxbd, bxb8, bxb1, bxb6, bxbe, bxb2, bxb3, bxb4, bxb7, bxbb, bxba, bxb9, bxbf)
 Digits(28)=Array(bxc0, bxc5, bxcc, bxcd, bxc8, bxc1, bxc6, bxce, bxc2, bxc3, bxc4, bxc7, bxcb, bxca, bxc9, bxcf)
 Digits(29)=Array(bxd0, bxd5, bxdc, bxdd, bxd8, bxd1, bxd6, bxde, bxd2, bxd3, bxd4, bxd7, bxdb, bxda, bxd9, bxdf)
 Digits(30)=Array(bxe0, bxe5, bxec, bxed, bxe8, bxe1, bxe6, bxee, bxe2, bxe3, bxe4, bxe7, bxeb, bxea, bxe9, bxef)
 Digits(31)=Array(bxf0, bxf5, bxfc, bxfd, bxf8, bxf1, bxf6, bxfe, bxf2, bxf3, bxf4, bxf7, bxfb, bxfa, bxf9, bxff)


Sub DisplayTimer()
  Dim ChgLED, ii, num, chg, stat, obj
  ChgLED=Controller.ChangedLEDs(&H00000000, &Hffffffff)
  If Not IsEmpty(ChgLED) Then
    If DesktopMode = True Then
    For ii=0 To UBound(chgLED)
      num=chgLED(ii,0):chg=chgLED(ii,1):stat=chgLED(ii,2)
      if (num < 32) then
      For Each obj In Digits(num)
        If chg And 1 Then obj.visible=stat And 1
        chg=chg\2:stat=stat\2
      Next
      else
        'For Each obj In Digits(num)
        ' If chg And 1 Then obj.state = stat And 1
        ' chg = chg\2 : stat = stat\2
        'Next
      end if
    Next
     end if
  End If
End Sub

'************GI Subs

 Dim GIActive,GIState
 GIActive=0:GIState=0

Dim GIActivebb,GIStatebb
 GIActivebb=0:GIStatebb=0

 Sub SolGION(enabled)
    FlippersEnabled = Enabled
    If enabled then
        GIActive=1:SolGI 1
    else
        GIActive=0:SolGI 0
        if leftflipper.startangle > leftflipper.endangle Then
            if leftflipper.currentangle < leftflipper.startangle then leftflipper.rotatetostart : leftflippersound 0 : end if
        elseif leftflipper.startangle < leftflipper.endangle Then
            if leftflipper.currentangle > leftflipper.startangle then leftflipper.rotatetostart : leftflippersound 0 : end If
        end If
        if rightflipper.startangle > rightflipper.endangle Then
            if rightflipper.currentangle < rightflipper.startangle then rightflipper.rotatetostart : rightflippersound 0 : end if
        elseif rightflipper.startangle < rightflipper.endangle Then
            if rightflipper.currentangle > rightflipper.startangle then rightflipper.rotatetostart : rightflippersound 0 : end If
        end If
    end if
 End Sub

 Sub SolbbGION(enabled)
  If enabled then
    GIActivebb=1:SolbbGI 1
  else
    GIActivebb=0:SolbbGI 0
  end if
 End Sub

 Sub SolGIBlink(enabled)
    If GIActive=1 then:SolGI Not enabled:end if
  If GIActivebb=1 then:SolbbGI Not enabled:end if
 End Sub

 'GI Lights

  Sub SolGI(Enabled)
    If enabled then
    Playsound "fx_relay_on"                             'ninuzzu - added relay click sound
'   Table1.ColorGradeImage = "ColorGrade_8"             'ninuzzu - added LUT color grade---->this will light the whole table when GI is on
    For each xx in aGiLights:xx.State = 1:next
    'If Sw28.IsDropped = 1 then: sw28l.State = 1: End if
    'If Sw27.IsDropped = 1 then: sw27l.State = 1: End if
    'If Sw26.IsDropped = 1 then: sw26l.State = 1: End if
    'If Sw25.IsDropped = 1 then: sw25l.State = 1: End if
    'If Sw39.IsDropped = 1 then: sw39l.State = 1: End if
    'If Sw38.IsDropped = 1 then: sw38l.State = 1: End if
    'If Sw37.IsDropped = 1 then: sw37l.State = 1: End if
    GIState=1
    TAFBackglass.image = "Cab_Backglass1"
    vrl54.visible = true
    VrCat.image = "cat1":VrWoman.image = "woman1"
    SetLamp 190, 0
    Primitive112.BlendDisableLighting = 10        'iaakki - lit these primitives
    Primitive56.BlendDisableLighting = 10
    else
    Playsound "fx_relay_off"                            'ninuzzu - added relay click sound
'   Table1.ColorGradeImage = "ColorGrade_1"         'ninuzzu - added LUT color grade---->this will darken the whole table when GI is off
    For each xx in aGiLights:xx.State = 0:next
    For each xx in TargetDropGi:xx.State = 0:next
    TAFBackglass.image = "Cab_Backglass0"
    vrl54.visible = false
    VrCat.image = "cat0":VrWoman.image = "woman0"
    GIState=0
    Primitive112.BlendDisableLighting = 0       'iaakki - unlit these primitives
    Primitive56.BlendDisableLighting = 0
    end if
 End Sub

  Sub SolbbGI(Enabled)
  If enabled then
  Playsound "fx_relay_on"               'ninuzzu - added relay click sound
  GIStatebb=1

  else
  Playsound "fx_relay_off"              'ninuzzu - added relay click sound
  GIStatebb=0

  end if
 End Sub

 '*******************************************************'
'               Fluppers Flasher Domes
'*******************************************************'
Dim TestFlashers, TableRef, FlasherLightIntensity, FlasherFlareIntensity, FlasherBloomIntensity, FlasherOffBrightness

                ' *********************************************************************
TestFlashers = 0        ' *** set this to 1 to check position of flasher object       ***
Set TableRef = Table1       ' *** change this, if your table has another name             ***
FlasherLightIntensity = 1   ' *** lower this, if the VPX lights are too bright (i.e. 0.1)   ***
FlasherFlareIntensity = 1   ' *** lower this, if the flares are too bright (i.e. 0.1)     ***
FlasherBloomIntensity = 0.7
FlasherOffBrightness = 1.5    ' *** brightness of the flasher dome when switched off (range 0-2)  ***
                ' *********************************************************************

Dim ObjLevel(20), objbase(20), objlit(20), objflasher(20), objbloom(20), objlight(20)
Dim tablewidth, tableheight : tablewidth = TableRef.width : tableheight = TableRef.height
'initialise the flasher color, you can only choose from "green", "red", "purple", "blue", "white" and "yellow"
'InitFlasher 1, "green" : InitFlasher 2, "red" : InitFlasher 3, "white"
'InitFlasher 4, "green" : InitFlasher 5, "red" : InitFlasher 6, "white"
'InitFlasher 7, "green" : InitFlasher 8, "red"
'InitFlasher 9, "green" : InitFlasher 10, "red" : InitFlasher 11, "white"

InitFlasher 2, "red" : InitFlasher 3, "yellow" : InitFlasher 4, "yellow"




' rotate the flasher with the command below (first argument = flasher nr, second argument = angle in degrees)
'RotateFlasher 4,17 : RotateFlasher 5,0 : RotateFlasher 6,90
'RotateFlasher 7,0 : RotateFlasher 8,0
'RotateFlasher 9,-45 : RotateFlasher 10,90 : RotateFlasher 11,90



Sub InitFlasher(nr, col)
  ' store all objects in an array for use in FlashFlasher subroutine
  Set objbase(nr) = Eval("Flasherbase" & nr) : Set objlit(nr) = Eval("Flasherlit" & nr)
  Set objflasher(nr) = Eval("Flasherflash" & nr) : Set objlight(nr) = Eval("Flasherlight" & nr)
  Set objbloom(nr) = Eval("Flasherbloom" & nr)
  ' If the flasher is parallel to the playfield, rotate the VPX flasher object for POV and place it at the correct height
  If objbase(nr).RotY = 0 Then
    objbase(nr).ObjRotZ =  atn( (tablewidth/2 - objbase(nr).x) / (objbase(nr).y - tableheight*1.1)) * 180 / 3.14159
    objflasher(nr).RotZ = objbase(nr).ObjRotZ : objflasher(nr).height = objbase(nr).z + 60
  End If
  ' set all effects to invisible and move the lit primitive at the same position and rotation as the base primitive
  objlight(nr).IntensityScale = 0 : objlit(nr).visible = 0 : objlit(nr).material = "Flashermaterial" & nr
  objlit(nr).RotX = objbase(nr).RotX : objlit(nr).RotY = objbase(nr).RotY : objlit(nr).RotZ = objbase(nr).RotZ
  objlit(nr).ObjRotX = objbase(nr).ObjRotX : objlit(nr).ObjRotY = objbase(nr).ObjRotY : objlit(nr).ObjRotZ = objbase(nr).ObjRotZ
  objlit(nr).x = objbase(nr).x : objlit(nr).y = objbase(nr).y : objlit(nr).z = objbase(nr).z
  objbase(nr).BlendDisableLighting = FlasherOffBrightness
  ' set the texture and color of all objects
  select case objbase(nr).image
    Case "dome2basewhite" : objbase(nr).image = "dome2base" & col : objlit(nr).image = "dome2lit" & col :
    Case "ronddomebasewhite" : objbase(nr).image = "ronddomebase" & col : objlit(nr).image = "ronddomelit" & col
    Case "domeearbasewhite" : objbase(nr).image = "domeearbase" & col : objlit(nr).image = "domeearlit" & col
  end select
  If TestFlashers = 0 Then objflasher(nr).imageA = "domeflashwhite" : objflasher(nr).visible = 0 : End If
  select case col
    Case "blue" :   objlight(nr).color = RGB(4,120,255) : objflasher(nr).color = RGB(200,255,255) : objlight(nr).intensity = 5000
    Case "green" :  objlight(nr).color = RGB(12,255,4) : objflasher(nr).color = RGB(12,255,4)
    Case "red" :    objlight(nr).color = RGB(255,32,4) : objflasher(nr).color = RGB(255,11,4)
    Case "purple" : objlight(nr).color = RGB(230,49,255) : objflasher(nr).color = RGB(255,64,255)
    Case "yellow" : objlight(nr).color = RGB(255,99,4) : objflasher(nr).color = RGB(255,177,4)
    Case "white" :  objlight(nr).color = RGB(255,240,150) : objflasher(nr).color = RGB(100,86,59)
  end select
  objlight(nr).colorfull = objlight(nr).color
  If TableRef.ShowDT and ObjFlasher(nr).RotX = -45 Then
    objflasher(nr).height = objflasher(nr).height - 20 * ObjFlasher(nr).y / tableheight
    ObjFlasher(nr).y = ObjFlasher(nr).y + 10
  End If
End Sub

Sub RotateFlasher(nr, angle) : angle = ((angle + 360 - objbase(nr).ObjRotZ) mod 180)/30 : objbase(nr).showframe(angle) : objlit(nr).showframe(angle) : End Sub

Sub FlashFlasher(nr)
  If not objflasher(nr).TimerEnabled Then objflasher(nr).TimerEnabled = True : objflasher(nr).visible = 1 : objbloom(nr).visible = 1 : objlit(nr).visible = 1 : End If
  objflasher(nr).opacity = 1000 *  FlasherFlareIntensity * ObjLevel(nr)^2.5
  objbloom(nr).opacity = 100 *  FlasherBloomIntensity * ObjLevel(nr)^2.5
  objlight(nr).IntensityScale = 0.5 * FlasherLightIntensity * ObjLevel(nr)^3
  objbase(nr).BlendDisableLighting =  FlasherOffBrightness + 10 * ObjLevel(nr)^3
  objlit(nr).BlendDisableLighting = 10 * ObjLevel(nr)^2
  UpdateMaterial "Flashermaterial" & nr,0,0,0,0,0,0,ObjLevel(nr),RGB(255,255,255),0,0,False,True,0,0,0,0
  ObjLevel(nr) = ObjLevel(nr) * 0.9 - 0.01
  If ObjLevel(nr) < 0 Then objflasher(nr).TimerEnabled = False : objflasher(nr).visible = 0 : objbloom(nr).visible = 0 : objlit(nr).visible = 0 : End If
End Sub

Sub FlasherFlash1_Timer() : FlashFlasher(1) : End Sub
Sub FlasherFlash2_Timer() : FlashFlasher(2) : End Sub
Sub FlasherFlash3_Timer() : FlashFlasher(3) : End Sub
Sub FlasherFlash4_Timer() : FlashFlasher(4) : End Sub
Sub FlasherFlash5_Timer() : FlashFlasher(5) : End Sub
Sub FlasherFlash6_Timer() : FlashFlasher(6) : End Sub
Sub FlasherFlash7_Timer() : FlashFlasher(7) : End Sub
Sub FlasherFlash8_Timer() : FlashFlasher(8) : End Sub
Sub FlasherFlash9_Timer() : FlashFlasher(9) : End Sub
Sub FlasherFlash10_Timer() : FlashFlasher(10) : End Sub
Sub FlasherFlash11_Timer() : FlashFlasher(11) : End Sub

'******************************
'BadCats Flasher Domes Config
'*******************************

' Upper Right Flasher'

Sub FlashRed2(flstate)
  If Flstate Then
      Objlevel(2) = 1 : FlasherFlash2_Timer
  End If
End Sub


Sub FlashYellow3(flstate)
  If Flstate Then
      Objlevel(3) = 1 : FlasherFlash3_Timer
  End If
End Sub

RotateFlasher 3, 0

'Upper Left Flasher

Sub FlashYellow4(flstate)
  If Flstate Then
      Objlevel(4) = 1 : FlasherFlash4_Timer
  End If
End Sub

RotateFlasher 4, 180

' Lower Right Flasher

RotateFlasher 2, -140

' Lower Left Flasher

'Sub FlashRed4(flstate)
'  If Flstate Then
'      Objlevel(4) = 4 : FlasherFlash4_Timer
 ' End If
'End Sub

'******************************************************
'        JP's VP10 Fading Lamps & Flashers
'  very reduced, mostly for rom activated flashers
' if you need to turn a light on or off then use:
'   LightState(lightnumber) = 0 or 1
'        Based on PD's Fading Light System
'******************************************************

Dim LightState(200), FlashSpeedUp(200), FlashSpeedDown(200), FlashMin(200), FlashMax(200), FlashLevel(200)
Dim FadeSpeedup, FadeSpeedDown, FadeMaxValue

FadeSpeedup = 100
FadeSpeedDown = 50
FadeMaxValue = 200

InitFlashers() ' turn off the lights and flashers and reset them to the default parameters

LampTimer.Interval = 50 'lamp fading speed
LampTimer.Enabled = 1

Sub LampTimer_timer()
    Dim chgLamp, x
    chgLamp = Controller.ChangedLamps
    If Not IsEmpty(chgLamp) Then
        For x = 0 To UBound(chgLamp)
            LightState(chgLamp(x, 0) ) = chgLamp(x, 1) 'light state as set by the rom
        Next
    End If
    ' Lights & Flashers
  LightX 1, l1
    FadeDisableLighting 1, p1
    LightX 2, l2
  FadeDisableLighting 2, p2
    LightX 3, l3
  FadeDisableLighting 3, p3
    LightX 4, l4
  FadeDisableLighting 4, p4
    LightX 5, l5
  FadeDisableLighting 5, p5
    LightX 6, l6
  FadeDisableLighting 6, p6
    LightX 7, l7
  FadeDisableLighting 7, p7
    LightX 8, l8
  FadeDisableLighting 8, p8
    Flash 9, l9
    Flash 10, l10
    Flash 11, l11
    Flash 12, l12
    Flash 13, l13
    LightX 14, l14
  FadeDisableLighting 14, p14
    LightX 15, l15
  FadeDisableLighting 15, p15
    LightX 16, l16
  FadeDisableLighting 16, p16
    Flash 17, l17
    Flash 18, l18
    Flash 19, l19

    LightX 21, l21
  FadeDisableLighting 21, p21
    LightX 22, l22
  FadeDisableLighting 22, p22
    LightX 23, l23
  FadeDisableLighting 23, p23
    LightX 24, l24
  FadeDisableLighting 24, p24
    LightX 25, l25
  FadeDisableLighting 25, p25
  LightX 26, l26
  FadeDisableLighting 26, p26
    LightX 27, l27
  FadeDisableLighting 27, p27
    LightX 28, l28
  FadeDisableLighting 28, p28
    LightX 29, l29
  FadeDisableLighting 29, p29
    LightX 30, l30 'dog house light I think
    FadeDisableLighting 30, p30
  LightX 31, l31
    LightX 33, l33
  FadeDisableLighting 33, p33
    LightX 34, l34
  FadeDisableLighting 34, p34
    LightX 35, l35
  FadeDisableLighting 35, p35
    LightX 36, l36
  FadeDisableLighting 36, p36
    LightX 37, l37
  FadeDisableLighting 37, p37
    LightX 38, l38
  FadeDisableLighting 38, p38
    LightX 39, l39
  FadeDisableLighting 39, p39
    LightX 40, l40
  FadeDisableLighting 40, p40
    LightX 41, l41
  FadeDisableLighting 41, p41
    LightX 42, l42
  FadeDisableLighting 42, p42
    LightX 43, l43
  FadeDisableLighting 43, p43
    LightX 44, l44
  FadeDisableLighting 44, p44
    LightX 45, l45
  FadeDisableLighting 45, p45
    LightX 46, l46
  FadeDisableLighting 46, p46
    LightX 47, l47
  FadeDisableLighting 47, p47
    LightX 48, l48
  FadeDisableLighting 48, p48
    LightX 49, l49
  FadeDisableLighting 49, p49
    LightX 50, l50
  FadeDisableLighting 50, p50
    LightX 51, l51
  FadeDisableLighting 51, p51
    LightX 52, l52
  FadeDisableLighting 52, p52
    LightX 53, l53
  FadeDisableLighting 53, p53
    'LightVrv 54, vrl54 'Lamp shed backglass
    LightVrv 55, vrl55  'bbq bg
    LightVrv 56, vrl56  'candle bg
    LightVrv 57, vrl57  '57-64 bg jackpot 1000000 - 8000000
    LightVrv 58, vrl58
    LightVrv 59, vrl59
    LightVrv 60, vrl60
    LightVrv 61, vrl61
    LightVrv 62, vrl62
    LightVrv 63, vrl63
    LightVrv 64, vrl64
    LightXm 125, f25a
    LightX 125, f25
  FadeDisableLighting 125, p125
    'LightXm 126, f26a
    'LightXm 126, f26
   ' Flash 126, f26b
    LightXm 127, f27
    lightXm 127, f27a
    Flash 127, f27b
    Flash 127, f27b2
    'LightXm 128, f28a ' f28, f28a, f28b are the old uppler left flasher before Flupper's dome update
    'LightXm 128, f28
  ' Flash 128, f28b
    LightXm 129, f29a
    LightXm 129, f29
    Flash 129, f29b
    LightXm 130, f30
   ' Flash 130, f30a
    ' LightXm 131, f31a
    ' LightXm 131, f31
    ' Flash 131, f31b
    LightXm 132, f32
'    Flash 132, f32a

'    Flash 190, SFWL
End Sub

'Sub FadeDisableLighting2(nr, a)
' Select Case LightState(nr)
'   Case 0
'     a.UserValue = a.UserValue - 0.25
'     If a.UserValue < 0 Then
'       a.UserValue = 0
'       LightState(nr) = -1
'     end If
'     a.BlendDisableLighting = 200 * a.UserValue 'brightness
'   Case 1
'     a.UserValue = a.UserValue + 0.5
'     If a.UserValue > 1 Then
'       a.UserValue = 1
'       LightState(nr) = -1
'     end If
'     a.BlendDisableLighting = 200 * a.UserValue 'brightness
' End Select
'End Sub

Sub FadeDisableLighting(nr, a)
  Select Case LightState(nr)
    Case 0
      If a.BlendDisableLighting <= (0 + FadeSpeedDown) Then
        a.BlendDisableLighting = 0
        LightState(nr) = -1
      Else
        a.BlendDisableLighting = a.BlendDisableLighting - FadeSpeedDown
      end If
    Case 1
      If a.BlendDisableLighting >= (FadeMaxValue - FadeSpeedup) Then
        a.BlendDisableLighting = FadeMaxValue
        LightState(nr) = -1
      Else
        a.BlendDisableLighting = a.BlendDisableLighting + FadeSpeedup
      end If
  End Select
End Sub

Sub SetLamp(nr, value)
    If value <> LightState(nr) Then
        LightState(nr) = value
    End If
End Sub

' div lamp subs

Sub InitFlashers()
    Dim x
    For x = 0 to 200
        LightState(x) = 0        ' light state: 0=off, 1=on, -1=no change (on or off)
        FlashSpeedUp(x) = 0.5    ' Fade Speed Up
        FlashSpeedDown(x) = 0.25 ' Fade Speed Down
        FlashMax(x) = 1          ' the maximum intensity when on, usually 1
        FlashMin(x) = 0          ' the minimum intensity when off, usually 0
        FlashLevel(x) = 0        ' the intensity/fading of the flashers
    Next
End Sub

' VPX Lights, just turn them on or off

Sub LightX(nr, object)
    Select Case LightState(nr)
        Case 0, 1:object.state = LightState(nr)':LightState(nr) = -1
    End Select
End Sub

Sub LightXm(nr, object) 'multiple lights
    Select Case LightState(nr)
        Case 0, 1:object.state = LightState(nr)
    End Select
End Sub

'vr bg lights turn on or off wit visibility primitive

Sub LightVrv(nr, object)
    Select Case LightState(nr)
        Case 0:object.visible = false:LightState(nr) = -1
        Case 1:object.visible = true:LightState(nr) = -1
    End Select
End Sub

Sub LightVrvm(nr, object) 'multiple lights
    Select Case LightState(nr)
        Case 0:object.visible = false
        Case 1:object.visible = true
    End Select
End Sub
'vr light swap image primitive
Sub LightVri(nr, object, a)
    Select Case LightState(nr)
        Case 0:object.image = a & 0:LightState(nr) = -1
        Case 1:object.image = a & 1:LightState(nr) = -1
    End Select
End Sub

Sub LightVrim(nr, object, a) 'multiple lights
    Select Case LightState(nr)
        Case 0:object.image = a & 0
        Case 1:object.image = a & 1
    End Select
End Sub

' VPX Flashers, changes the intensity

Sub Flash(nr, object)
    Select Case LightState(nr)
        Case 0 'off
            FlashLevel(nr) = FlashLevel(nr) - FlashSpeedDown(nr)
            If FlashLevel(nr) < FlashMin(nr) Then
                FlashLevel(nr) = FlashMin(nr)
                LightState(nr) = -1 'completely off, so stop the fading loop
            End if
            Object.IntensityScale = FlashLevel(nr)
        Case 1 ' on
            FlashLevel(nr) = FlashLevel(nr) + FlashSpeedUp(nr)
            If FlashLevel(nr) > FlashMax(nr) Then
                FlashLevel(nr) = FlashMax(nr)
                LightState(nr) = -1 'completely on, so stop the fading loop
            End if
            Object.IntensityScale = FlashLevel(nr)
    End Select
End Sub

Sub Flashm(nr, object) 'multiple flashers, it just sets the intensity
    Object.IntensityScale = FlashLevel(nr)
End Sub

' *********************************************************************
'                      Supporting Ball & Sound Functions
' *********************************************************************

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
    Vol = Csng(BallVel(ball) ^2 / 500)
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

Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
    Pitch = BallVel(ball) * 20
End Function

Function BallVel(ball) 'Calculates the ball speed
    BallVel = INT(SQR((ball.VelX ^2) + (ball.VelY ^2) ) )
End Function

' *******************************************************************************************************
' Positional Sound Playback Functions by DJRobX
' PlaySoundAtVol sound, 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 1, AudioFade(ActiveBall)
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

Sub PlaySoundAtVol(sound, tableobj, Volum)
  PlaySound sound, 1, Volum, Pan(tableobj), 0,0,0, 1, AudioFade(tableobj)
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
'                    More Supporting Ball & Sound Functions :)
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
'      JP's VP10 Rolling Sounds
'*****************************************

'Const tnob = 5 ' total number of balls in this table is 4, but always use a higher number here because of the timing

Const tnob = 1 'ninuzzu - why 5 balls? Bad Cats has only one ball
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
            If BOT(b).z < 30 Then
                ballpitch = Pitch(BOT(b) )
            Else
                ballpitch = Pitch(BOT(b) ) * 100
            End If
            rolling(b) = True
            PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) )*0.01, Pan(BOT(b) ), 0, ballpitch, 1, 0
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
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 500, Pan(ball1), 0, Pitch(ball1), 0, 0
End Sub

'*****************************************
'   Ball Shadow
'*****************************************

Dim BallShadow
BallShadow = Array (BallShadow1)                                                                        'ninuzzu - let's create an array of primitives, the number of primitives is equal to tnob
Dim ShadowSFW
ShadowSFW = 0

Sub shadowTrig_Hit: ShadowSFW = 1: End Sub                                                              'ninuzzu- so in this case only one primitive, for 3 ball it will be BallShadow = Array (BallShadow1,BallShadow2,BallShadow3)
Sub shadowTrig_UnHit: ShadowSFW = 0: End Sub

Sub BallShadowUpdate()
    Dim BOT, b
    BOT = GetBalls                                                                                      'ninuzzu- this will return an array , the balls array, this is updated in real time

    ' render the shadow for each ball
    For b = 0 to UBound(BOT)                                                                            'ninuzzu - now let's link the ball array with the array of primitives; so for each ball in the array, do this
        If BOT(b).X < Table1.Width/2 Then
            BallShadow(b).X = ((BOT(b).X) - (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/7)) + 10      'ninuzzu - the shadow array will move left or right depending on the ball X position in the table
        Else
            BallShadow(b).X = ((BOT(b).X) + (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/7)) - 10
        End If
        BallShadow(b).Y = BOT(b).Y + 20                                                                 'ninuzzu - the shadow Y is at ball Y + 20 units lower
        BallShadow(b).Z = 1                                                                             'ninuzzu - the shadow Z is 1

        If (BOT(b).Z > 20 and ShadowSFW = 0)  Then                                                                          'ninuzzu - if the ball is falling through a hole, e.g. a subway, the shadow is not visible.
            BallShadow(b).visible = 1
        Else
            BallShadow(b).visible = 0
        End If
    Next
End Sub

'******************************
' Diverse Collection Hit Sounds
'******************************

Sub aMetal_Wires_Hit(idx):PlaySound "fx_metalhit", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
Sub aRubber_Bands_Hit(idx):PlaySound "fx_rubber_band", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
Sub aRubber_Posts_Hit(idx):PlaySound "fx_rubber", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0: dPosts():End Sub
Sub aRubber_Sleeves_Hit(idx):PlaySound "fx_postrubber", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0: dSleeves(): End Sub
Sub aPlastics_Hit(idx):PlaySound "fx_PlasticHit", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
Sub aGates_Hit(idx):PlaySound "fx_Gate", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
Sub aWoods_Hit(idx):PlaySound "fx_Woodhit", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub


'**************************
'   Option Setup
'**************************

Sub TableOptions()
  if CabinetMode = 1 then
    Korpus.Size_y=2
    Siderails.visible=0
    Lockdownbar.visible=0
  Else
    Korpus.Size_y=1
    Siderails.visible=1
    Lockdownbar.visible=1
  end If
End Sub


DIM VRThings
If VRRoom > 0 Then
  glass.visible=0
  Siderails.visible=1
  Lockdownbar.visible=1
  If VRRoom = 1 Then
    for each VRThings in VRCab:VRThings.visible=1:Next
    for each VRThings in VRBG:VRThings.visible=1:Next
    for each VRThings in VRMin:VRThings.visible=0:Next
    for each VRThings in VRStuff:VRThings.visible=1:Next
    for each VRThings in VRLights:VRThings.State=1:Next
    FlasherforLights.visible=1
  End If

  If VRRoom = 2 Then
    for each VRThings in VRCab:VRThings.visible=1:Next
    for each VRThings in VRBG:VRThings.visible=1:Next
    for each VRThings in VRMin:VRThings.visible=1:Next
    for each VRThings in VRStuff:VRThings.visible=0:Next
    for each VRThings in VRLights:VRThings.State=0:Next
  End If

  If VRRoom = 3 Then
    for each VRThings in VRCab:VRThings.visible=0:Next
    for each VRThings in VRBG:VRThings.visible=1:Next
    for each VRThings in VRMin:VRThings.visible=0:Next
    for each VRThings in VRStuff:VRThings.visible=0:Next
    for each VRThings in VRLights:VRThings.State=0:Next
    for each VRThings in VRCab:VRThings.visible = 0:Next
  End If

  Else
    for each VRThings in VRCab:VRThings.visible=0:Next
    for each VRThings in VRBG:VRThings.visible=0:Next
    for each VRThings in VRMin:VRThings.visible=0:Next
    for each VRThings in VRStuff:VRThings.visible=0:Next
    for each VRThings in VRLights:VRThings.State=0:Next
    FlasherforLights.visible=0
    if DesktopMode then
      Siderails.visible=1
      Lockdownbar.visible=1
    else
      Siderails.visible=0
      Lockdownbar.visible=0
    End If
End if

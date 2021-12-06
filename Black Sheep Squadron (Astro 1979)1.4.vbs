'Black Sheep Squadron (Astro 1979)

'Version 1.2

' Thalamus 2018-07-19
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Changed UseSolenoids=1 to 2
' Thalamus 2018-08-09 : Improved directional sounds
' !! NOTE : Table not verified yet !!

Option Explicit
Randomize

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
Const VolWood   = 1    ' Woods volume.
Const VolKick   = 1    ' Kicker volume.
Const VolSpin   = 1.5  ' Spinners volume.
Const VolFlip   = 1    ' Flipper volume.


On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Const cGameName="blkshpsq", UseSolenoids=2,UseLamps=1,UseGI=0,SSolenoidOn="SolOn",SSolenoidOff="SolOff",SFlipperOn="fx_Flipperup",SFlipperOff="fx_Flipperdown", SCoin="Coin",cCredits=""
Const BallMass = 1.7

LoadVPM "01550000", "Bally.vbs", 3.26
Dim DesktopMode: DesktopMode = Table1.ShowDT

'*********
'Solenoids
'*********
 Solcallback(2) = "vpmSolSound SoundFX(""Chime10"",DOFChimes),"
 Solcallback(3) = "vpmSolSound SoundFX(""Chime100"",DOFChimes),"
 Solcallback(4) = "vpmSolSound SoundFX(""Chime1000"",DOFChimes),"
 Solcallback(5) = "vpmSolSound SoundFX(""ChimeExtra"",DOFChimes),"
 SolCallback(6) = "vpmSolSound SoundFX(""knocker"",DOFKnocker),"
 SolCallback(7) = "bsTrough.SolOut"
 SolCallBack(10) = "SolSaucer"    '"bsSaucer.SolOut"
 SolCallback(19) = "vpmNudge.SolGameOn"

'**************
' Flipper Subs
'**************

 SolCallback(sLRFlipper) = "SolRFlipper"
 SolCallback(sLLFlipper) = "SolLFlipper"

 Sub SolLFlipper(Enabled)
     If Enabled Then
         PlaySoundAtVol SoundFX("fx_Flipperup",DOFFlippers),LeftFlipper, VolFlip:LeftFlipper.RotateToEnd:PlaySoundAtVol "buzz", LeftFlipper, VolFlip
     Else
         PlaySoundAtVol SoundFX("fx_Flipperdown",DOFFlippers),LeftFlipper, VolFlip:LeftFlipper.RotateToStart:StopSound "buzz"
     End If
  End Sub

 Sub SolRFlipper(Enabled)
     If Enabled Then
         PlaySoundAtVol SoundFX("fx_Flipperup",DOFFlippers), RightFlipper, VolFlip:RightFlipper.RotateToEnd:PlaysoundAtVol "buzz1", RightFlipper, VolFlip
     Else
         PlaySoundAtVol SoundFX("fx_Flipperdown",DOFFlippers), RightFlipper, VolFlip:RightFlipper.RotateToStart:StopSound "buzz1"
     End If
 End Sub

Sub LeftFlipper_Collide(parm)
    RandomSoundFlipper()
End Sub

Sub RightFlipper_Collide(parm)
    RandomSoundFlipper()
End Sub

Sub RandomSoundFlipper()
    Select Case Int(Rnd*3)+1
        Case 1 : PlaySound "flip_hit_1", 0, Vol(ActiveBall)*VolRH, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
        Case 2 : PlaySound "flip_hit_2", 0, Vol(ActiveBall)*VolRH, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
        Case 3 : PlaySound "flip_hit_3", 0, Vol(ActiveBall)*VolRH, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
    End Select
End Sub

'**************
' GI Lights On
'**************

dim xx
For each xx in GI:xx.State = 1: Next

'Initiate Table
'**********************************************************************************************************

Dim bsTrough,bsSaucer

Sub Table1_Init
    vpmInit Me
    On Error Resume Next
        With Controller
        .GameName = cGameName
        If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
        .SplashInfoLine="Black Sheep Squadron (Astro 1979)"
        .HandleMechanics=0
        .HandleKeyboard=0
        .ShowDMDOnly=1
        .ShowFrame=0
        .ShowTitle=0
        .Hidden = 1
        If Err Then MsgBox Err.Description
    End With
    On Error Goto 0
        ' Controller.Dip(0) = (0*1 + 0*2 + 0*4 + 0*8 + 0*16 + 1*32 + 1*64 + 0*128) '01-08
        ' Controller.Dip(1) = (0*1 + 0*2 + 0*4 + 0*8 + 0*16 + 1*32 + 0*64 + 1*128) '09-16
        ' Controller.Dip(2) = (1*1 + 0*2 + 1*4 + 1*8 + 1*16 + 0*32 + 0*64 + 1*128) '17-24
        ' Controller.Dip(3) = (1*1 + 0*2 + 0*4 + 0*8 + 0*16 + 0*32 + 0*64 + 0*128) '25-32

        Controller.SolMask(0)=0
        vpmTimer.AddTimer 2000,"Controller.SolMask(0)=&Hffffffff'" 'ignore all solenoids - then add the timer to renable all the solenoids after 2 seconds
        Controller.Run
    If Err Then MsgBox Err.Description
    On Error Goto 0

    PinMAMETimer.Interval=PinMAMEInterval
    PinMAMETimer.Enabled=1


     ' Nudging
     vpmNudge.TiltSwitch = 7
     vpmNudge.Sensitivity = 5
     vpmNudge.TiltObj = Array(Bumper1, Bumper2, Bumper3, LeftSlingshot, RightSlingshot)

     ' Trough
     Set bsTrough = New cvpmBallStack
     With bsTrough
         .InitSw 0, 8, 0, 0, 0, 0, 0, 0
         .InitKick BallRelease, 90, 7
         .InitEntrySnd "Solenoid", "Solenoid"
         .InitExitSnd SoundFX("ballrelease",DOFContactors), SoundFX("Solenoid",DOFContactors)
         .Balls = 1
     End With

     ' Saucer
    Set bsSaucer=New cvpmBallStack
    bsSaucer.InitSaucer Kicker1,13,131,10
    bsSaucer.InitExitSnd SoundFX("metalhit2",DOFContactors),SoundFX("popper_ball",DOFContactors)

End Sub



Sub table1_Paused:Controller.Pause = 1:End Sub
Sub table1_unPaused:Controller.Pause = 0:End Sub

'**********************************************************************************************************
'Key Handling
'**********************************************************************************************************
 Sub Table1_KeyDown(ByVal keycode)
    If vpmKeyDown(KeyCode) Then Exit Sub

    If keycode=PlungerKey Then Plunger.Pullback:playsoundAtVol"plungerpull",Plunger, 1

    If keycode = LeftTiltKey Then LeftNudge 80, 1.2, 20:PlaySound SoundFX("nudge_left",0)
    If keycode = RightTiltKey Then RightNudge 280, 1.2, 20:PlaySound SoundFX("nudge_right",0)
    If keycode = CenterTiltKey Then CenterNudge 0, 1.6, 25:PlaySound SoundFX("nudge_forward",0)
End Sub

Sub Table1_KeyUp(ByVal keycode)
    If vpmKeyUp(KeyCode) Then Exit Sub
    If keycode=PlungerKey Then Plunger.Fire:playsoundAtVol"plunger", Plunger, 1

End Sub


'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep

Sub RightSlingShot_Slingshot
    PlaySoundAtVol SoundFX("right_slingshot",DOFContactors), sling1, 1
    vpmTimer.PulseSw 36
    RightSling.Visible = 0
    RightSling1.Visible = 1
    sling1.TransZ = -20
    RStep = 0
    RightSlingShot.TimerEnabled = 1

End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RightSLing1.Visible = 0:RightSLing2.Visible = 1:sling1.TransZ = -10
        Case 4:RightSLing2.Visible = 0:RightSLing.Visible = 1:sling1.TransZ = 0:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
    PlaySoundAtVol SoundFX("left_slingshot",DOFContactors), sling2, 1
    vpmTimer.PulseSw 37
    LeftSling.Visible = 0
    LeftSling1.Visible = 1
    sling2.TransZ = -20
    LStep = 0
    LeftSlingShot.TimerEnabled = 1

End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LeftSLing1.Visible = 0:LeftSLing2.Visible = 1:sling2.TransZ = -10
        Case 4:LeftSLing2.Visible = 0:LeftSLing.Visible = 1:sling2.TransZ = 0:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub


'**********Rubber Compression
' RUstep, RLstep, and LLstep  are the variables that increment the animation
'****************
Dim RUStep, RLstep, LLstep, LUstep, LDstep

Sub RubberRL_Hit
    'Rubber Animation
    RubberRight.Visible = 0
    RubberRightLower.Visible = 1
    RLstep = 0
    RubberRL.TimerEnabled = 1
End Sub


Sub RubberRL_Timer
    Select Case RLstep
        Case 1:RubberRightLower.Visible = 0:RubberRight.Visible = 1:RubberRL.TimerEnabled = 0
    End Select
    RLStep = RLStep + 1
End Sub

Sub RubberLU_Hit
    'Rubber Animation
    RubberLeftUpper.Visible = 0
    RubberLeftUpper1.Visible = 1
    LUstep = 0
    RubberLU.TimerEnabled = 1
End Sub


Sub RubberLU_Timer
    Select Case LUstep
        Case 1:RubberLeftUpper1.Visible = 0:RubberLeftUpper.Visible = 1:RubberLU.TimerEnabled = 0
    End Select
    LUStep = LUStep + 1
End Sub



'*********
' Switches
'*********

'Drain hole
Sub Drain_Hit:bsTrough.AddBall Me : PlaySoundAtVol "drain", drain,1 :End Sub

'Kicker
Sub Kicker1_hit():bsSaucer.AddBall 0:playsoundAtVol "popper_ball", kicker1, VolKick: End Sub

Sub SolSaucer(enabled)
    If Enabled Then
        kicker1.kick 131,10
        Controller.Switch(12) = False
        PlaySoundAtVol "popper_ball", kicker1, VolKick
        kicker1.uservalue=1
        kicker1.timerenabled=1
        Pkickarm.rotz=15
    end if
End Sub

Sub kicker1_timer
    select case kicker1.uservalue
      case 2:
        Pkickarm.rotz=0
        me.timerenabled=0
    end Select
    kicker1.uservalue=kicker1.uservalue+1
End Sub

'10 point Rubbers
 Sub sw17b_Hit:vpmTimer.PulseSw 17
    vpmTimer.PulseSw 17
    'Rubber Animation
    RubberRight.Visible = 0
    RubberRightUpper.Visible = 1
    RUstep = 0
    sw17b.TimerEnabled = 1
End Sub

Sub sw17a_Hit:vpmTimer.PulseSw 17
    vpmTimer.PulseSw 17
    'Rubber Animation
    RubberLeft.Visible = 0
    RubberLeftUpper.Visible = 1
    RUstep = 0
    sw17a.TimerEnabled = 1
End Sub

' Bumpers
Dim bump1,bump2,bump3

Sub Bumper1_Hit:vpmTimer.PulseSw 38:PlaySoundAtVol SoundFX("fx_bumper1",DOFContactors), Bumper1, VolBump:bump1 = 1:Me.TimerEnabled = 1::End Sub
    Sub Bumper1_timer()
            Select Case bump1
               Case 1:Ring1.z = 5:bump1 = 2
               Case 2:Ring1.z = -5:bump1 = 3
               Case 3:Ring1.z = -15:bump1 = 4
               Case 4:Ring1.z = -15:bump1 = 5
               Case 5:Ring1.z = -5:bump1 = 6
               Case 6:Ring1.z = 5:bump1 = 7
               Case 7:Ring1.z = 5:Me.TimerEnabled = 0
            End Select
        End Sub

Sub Bumper2_Hit:vpmTimer.PulseSw 39:PlaySoundAtVol SoundFX("fx_bumper2",DOFContactors), Bumper2, VolBump: bump2 = 1:Me.TimerEnabled = 1::End Sub
    Sub Bumper2_timer()
                Select Case bump2
                Case 1:Ring2.z = 5:bump2 = 2
                Case 2:Ring2.z = -5:bump2 = 3
                Case 3:Ring2.z = -15:bump2 = 4
                Case 4:Ring2.z = -15:bump2 = 5
                Case 5:Ring2.z = -5:bump2 = 6
                Case 6:Ring2.z = 5:bump2 = 7
                Case 7:Ring2.z = 5:Me.TimerEnabled = 0
            End Select
        End Sub

Sub Bumper3_Hit:vpmTimer.PulseSw 40:PlaySoundAtVol SoundFX("fx_bumper3",DOFContactors), Bumper3, VolBump:bump3 = 1:Me.TimerEnabled = 1: End Sub
Sub Bumper3_timer()
           Select Case bump3
              Case 1:Ring3.z = 5:bump3 = 2
              Case 2:Ring3.z = -5:bump3 = 3
              Case 3:Ring3.z = -15:bump3 = 4
              Case 4:Ring3.z = -15:bump3 = 5
              Case 5:Ring3.z = -5:bump3 = 6
              Case 6:Ring3.z = 5:bump3 = 7
              Case 7:Ring3.z = 5:Me.TimerEnabled = 0
           End Select
       End Sub


' Rollovers
Sub sw34_Hit:Controller.Switch(34) = 1:PlaySoundAtVol "sensor", ActiveBall, 1:PlaySoundAtVol "outlane", ActiveBall, 1:End Sub 'L outlane
Sub sw34_UnHit:Controller.Switch(34) = 0:End Sub

Sub sw33_Hit:Controller.Switch(33) = 1:PlaySoundAtVol "sensor", ActiveBall, 1:PlaySoundAtVol "outlane", ActiveBall, 1:End Sub 'R outlane
Sub sw33_UnHit:Controller.Switch(33) = 0:End Sub

Sub sw21_Hit:Controller.Switch(21) = 1:PlaySoundAtVol "sensor", ActiveBall, 1:End Sub 'Lane R
Sub sw21_UnHit:Controller.Switch(21) = 0:End Sub

Sub sw22_Hit:Controller.Switch(22) = 1:PlaySoundAtVol "sensor", ActiveBall, 1:End Sub 'Lane D
Sub sw22_UnHit:Controller.Switch(22) = 0:End Sub

Sub sw23_Hit:Controller.Switch(23) = 1:PlaySoundAtVol "sensor", ActiveBall, 1:End Sub 'Lane A
Sub sw23_UnHit:Controller.Switch(23) = 0:End Sub

Sub sw24_Hit:Controller.Switch(24) = 1:PlaySoundAtVol "sensor", ActiveBall, 1:End Sub 'Lane U
Sub sw24_UnHit:Controller.Switch(24) = 0:End Sub

Sub sw26_Hit:Controller.Switch(26) = 1:PlaySoundAtVol "sensor", ActiveBall, 1:End Sub 'L inlane
Sub sw26_UnHit:Controller.Switch(26) = 0:End Sub

Sub sw25_Hit:Controller.Switch(25) = 1:PlaySoundAtVol "sensor", ActiveBall, 1:End Sub 'R inlane
Sub sw25_UnHit:Controller.Switch(25) = 0:End Sub

Sub sw12a_Hit:Controller.Switch(12) = 1:PlaySoundAtVol "sensor", ActiveBall, 1:End Sub 'Bottom Rollover
Sub sw12a_UnHit:Controller.Switch(12) = 0:End Sub

Sub sw12_Hit:Controller.Switch(12) = 1:PlaySoundAtVol "sensor", ActiveBall, 1:End Sub 'Top Rollover
Sub sw12_UnHit:Controller.Switch(12) = 0:End Sub

Sub sw1_Hit:Controller.Switch(1) = 1:PlaySoundAtVol "sensor", ActiveBall, 1:End Sub 'Chute rollover
Sub sw1_UnHit:Controller.Switch(1) = 0:End Sub

' Targets
Sub sw29_Hit:vpmTimer.PulseSw 29:PlaySoundAtVol SoundFX("target",DOFTargets), sw29, VolTarg:End Sub  'Target N
Sub sw30_Hit:vpmTimer.PulseSw 30:PlaySoundAtVol SoundFX("target",DOFTargets), sw30, VolTarg:End Sub  'Target O
Sub sw31_Hit:vpmTimer.PulseSw 31:PlaySoundAtVol SoundFX("target",DOFTargets), sw31, VolTarg:End Sub  'Target Q
Sub sw32_Hit:vpmTimer.PulseSw 32:PlaySoundAtVol SoundFX("target",DOFTargets), sw32, VolTarg:End Sub  'Target S

'  Spinner1
Sub SPinner1_Spin():vpmTimer.PulseSw(2) : playsoundAtVol"fx_spinner", Spinner1, VolSpin : End Sub


'**************
' Edit Dips
'**************
    '******About the DIP switches******
'The astro rom and bally.vbs are not quite 100% compatible
'some of the switches in the manual are correct some are not in earlier versions
'the third light at the kicker hole would not cycle and you could not add more than 2 players
'Dip 24 and 25 "ON" solve these issues and must be on for the game to function properly
'The adjustable Dips are:
'Dip 7 balls per game  on=3 off=5
'Dip 8 The song "Halls of Montezuma" on the chimes will play  off=Only at startup  on= With each player add
'Dip 16 Bumper and extra ball outlane lights alternate  on= alternate  off= stay on
'Dip 22 Bonus countdown will cycle through the spots x's how many multipliers or count x's at each spot On=over x's mult  off= count mult at each spot

Sub EditDips
Dim vpmDips:Set vpmDips=New cvpmDips 'initializes vpmDips to the cvpmDips class to allow creation of a window option menu
With vpmDips
.AddForm 700,400,"Black Sheep - DIP switches"
.AddChk 2,2,190,Array("dip  1 OFF",&H00000001)'dip 1    "1-5 Control coin slot 1 all off is 1 for 1"
.AddChk 2,20,190,Array("dip 2 OFF",&H00000002)'dip 2
.AddChk 2,40,190,Array("dip 3 OFF",&H00000004)'dip 3
.AddChk 2,60,190,Array("dip 4 OFF",&H00000004)'dip 4
.AddChk 2,80,190,Array("dip 5 OFF",&H00000010)'dip 5
.AddChk 2,100,190,Array("dip 6 ON",&H00000020)'dip 6 to give a credit for beating the high game
.AddChk 2,120,190,Array("*dip 7 BPG ON=5 OFF=3",&H00000040)'dip 7 off 3 on 5 balls per game
.AddChk 2,140,190,Array("*dip 8 Song Off=only at begining",&H00000080)'dip 8 halls of montezuma plays at begining if off or on every player add if on
.AddChk 2,160,190,Array("dip 9 OFF",&H00000100)'dip 9   "9-13 control coin slot 2 all of is 3 for 2"
.AddChk 2,180,190,Array("dip 10 OFF",&H00000200)'dip 10
.AddChk 2,200,190,Array("dip 11 OFF",&H00000400)'dip 11
.AddChk 2,220,190,Array("dip 12 OFF",&H00000800)'dip 12
.AddChk 2,240,190,Array("dip 13 OFF",&H00001000)'dip 13
.AddChk 2,260,190,Array("dip 14 ON",&H00002000)'dip 14 & 15 current settings give 1 game for HS
.AddChk 2,280,190,Array("dip 15 OFF",&H00004000)'dip 15
.AddChk 2,300,190,Array("*dip 16 Bumper lights ON=Alternate",32768)'dip 16 for bumper and in lane lights to alternate or stay on
.AddChk 200,2,190,Array("dip 17 ON",&H00010000)'dip 17          17,18,19 control max credits this setup is for 30 max
.AddChk 200,20,190,Array("dip 18 OFF",&H00020000)'dip 18
.AddChk 200,40,190,Array("dip 19 ON",&H00040000)'dip 19
.AddChk 200,60,190,Array("dip 20 ON credit display",&H00080000)'dip 20
.AddChk 200,80,190,Array("dip 21 ON match Function",&H00100000)'dip 21
.AddChk 200,100,190,Array("*dip 22 OFF bns countdown",&H00200000)'dip 22 if on bonus countdown will go through all spots x multiplier - off will count multiplier at each spot
.AddChk 200,120,190,Array("dip 23 OFF",&H00400000)'dip 23  unkown
.AddChk 200,140,190,Array("dip 24 ON For l40 to function",&H00800000)'dip 24 must be on for l40 at the kicker to function otherwise it does not cycle through the "light extra ball lanes" lamp
.AddChk 200,160,190,Array("dip 25 ON for 4 player",&H01000000)'dip 25 must be on for the 4 player to work
.AddChk 200,180,190,Array("dip 26 OFF",&H02000000)'dip 26  26 - 30 are not used by the atro rom but may trigger something inthe bally.vbs
.AddChk 200,200,190,Array("dip 27 OFF",&H04000000)'dip 27
.AddChk 200,220,190,Array("dip 28 OFF",&H08000000)'dip 28
.AddChk 200,240,190,Array("dip 29 OFF",&H10000000)'dip 29
.AddChk 200,260,190,Array("dip 30 OFF",&H20000000)'dip 30
.AddChk 200,280,190,Array("dip 31 OFF",&H40000000)'dip 31  31 and 32 off will give 10,000 points for hitting special
.AddChk 200,300,190,Array("dip 32 OFF",&H80000000)'dip 32
.Addlabel 50,320,340,20,"Rom is not fully Compatible with the Bally.vbs"
.Addlabel 50,340,340,20, "the * dips are switchable"
.AddLabel 50,400,340,20,"After hitting OK, press F3 to reset game with new settings."
.ViewDips
End With
End Sub
Set vpmShowDips=GetRef("EditDips") ' This sets the F6 routine for vpmShowDips to run the subroutine "EditDips" rather than the standard dip menu in the VBS file


'********
'LIGHTS
'********

Set Lights(1) = l1 '1k bouns
Set Lights(2) = l2 '5k bonus
Set Lights(3) = l3 '9k bonus
Set Lights(4) = l4 'Target S
Set Lights(5) = l5 'Lane U
Lights(6) = Array(l6,l6a,l6b,l6c,l6d,l6e) 'Top & Bottom bumper lights
Lights(7) = Array(l7,l7a,l7b) ' middle bumper light
Set Lights(8) = l8 'Lite Spinner at kicker
Set Lights(9) = l9 '2x bounus enable
Set Lights(10) = l10 '2x bonus
Set Lights(11) = l11 'Shoot again
Set Lights(12) = l12 'BG Same player shoots again
Set Lights(13) = l13 'BG Ball in Play
Set Lights(14) = l14 'BG Can Play 1
Set Lights(15) = l15 'BG Player Up 1
Set Lights(17) = l17 '2k bonus
Set Lights(18) = l18 '6k bonus
Set Lights(19) = l19 '10k bonus
Set Lights(20) = l20 'Target Q
Set Lights(21) = l21 'Lane A
Set Lights(22) = l22 'L inlane
Set Lights(23) = l23 'R inlane
Set Lights(24) = l24 '50000 at kicker
Set Lights(25) = l25 '3x bonus enable
Set Lights(26) = l26 '3x bonus
Set Lights(27) = l27 'BG Match
Set Lights(29) = l29 'BG High Score to Date
Set Lights(30) = l30 'BG Can Play 2
Set Lights(31) = l31 'BG Player Up 2
Set Lights(33) = l33 '3k bonus
Set Lights(34) = l34 '7k bonus
Set Lights(36) = l36 'Target O
Set Lights(37) = l37 'Lane D
Set Lights(38) = l38 'top rollover
Set Lights(39) = l39 'Bottom rollover
Set Lights(40) = l40 'unknown should be lite extra ball lanes at kicker???
Set Lights(41) = l41 '5x bonus enable
Set Lights(42) = l42 '5x Bonus
Set Lights(45) = l45 'BG GAME OVER
Set Lights(46) = l46 'BG can play 3
Set Lights(47) = l47 'BG Player up 3
Set Lights(49) = l49 '4k bonus
Set Lights(50) = l50 '8k bonus
Set Lights(52) = l52 'Target N
Set Lights(53) = l53 'Slot R
'Set Lights(54) = l54 'on constant in game
'Set Lights(55) = l55 'on constant in game
Set Lights(57) = l57 'Special Light
Set Lights(60) = l60  'Spinner light
Set Lights(61) = l61 'BG TILT
Set Lights(62) = l62 'BG can play 4
Set Lights(63) = l63 'BG Player up 4


'******************************************************
'               RealTime Updates
'******************************************************

dim defaultEOS,EOSAngle,EOSTorque
defaulteos = leftflipper.eostorque
EOSAngle = 3
EOSTorque = 0.9

Sub GameTimer_Timer()
    RollingSoundUpdate


    If LeftFlipper.CurrentAngle < LeftFlipper.EndAngle + EOSAngle Then
        LeftFlipper.eostorque = EOSTorque
    Else
        LeftFlipper.eostorque = defaultEOS
    End If

    If RightFlipper.CurrentAngle > RightFlipper.EndAngle - EOSAngle Then
        RightFlipper.eostorque = EOSTorque
    Else
        RightFlipper.eostorque = defaultEOS
    End If

    FlipperLSh.RotZ = LeftFlipper.currentangle
    FlipperRSh.RotZ = RightFlipper.currentangle

End Sub

'*********** BALL SHADOW *********************************
Dim BallShadow
BallShadow = Array (BallShadow1,BallShadow2,BallShadow3,BallShadow4,BallShadow5)

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
            BallShadow(b).X = ((BOT(b).X) - (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/7)) + 10
        Else
            BallShadow(b).X = ((BOT(b).X) + (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/7)) - 10
        End If
        ballShadow(b).Y = BOT(b).Y + 20
        If BOT(b).Z > 20 Then
            BallShadow(b).visible = 1
        Else
            BallShadow(b).visible = 0
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
        PlaySound "fx_rubber2", 0, Vol(ActiveBall)*VolRH, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    End if
    If finalspeed >= 6 AND finalspeed <= 20 then
        RandomSoundRubber()
    End If
End Sub


Sub Rubbers_Collidable_Hit(idx)
    dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
    If finalspeed > 20 then
        PlaySound "fx_rubber2", 0, Vol(ActiveBall)*VolRH, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    End if
    If finalspeed >= 6 AND finalspeed <= 20 then
        RandomSoundRubber()
    End If
End Sub

Sub Rubbers_Collidable_Wall_Hit(idx)
    dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
    If finalspeed > 20 then
        PlaySound "fx_rubber2", 0, Vol(ActiveBall)*VolRH, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    End if
    If finalspeed >= 6 AND finalspeed <= 20 then
        RandomSoundRubber()
    End If
End Sub

Sub Posts_Hit(idx)
    dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
    If finalspeed > 16 then
        PlaySound "fx_rubber2", 0, Vol(ActiveBall)*VolPo, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    End if
    If finalspeed >= 6 AND finalspeed <= 16 then
        RandomSoundRubber()
    End If
End Sub

Sub RandomSoundRubber()
    Select Case Int(Rnd*2)+1
        Case 1 : PlaySound "rubber_hit_1", 0, Vol(ActiveBall)*VolRH, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
        Case 2 : PlaySound "rubber_hit_2", 0, Vol(ActiveBall)*VolRH, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
        Case 3 : PlaySound "rubber_hit_3", 0, Vol(ActiveBall)*VolRH, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    End Select
End Sub

Sub Bumpers_Hit(idx)
    Select Case Int(Rnd*4)+1
        Case 1 : PlaySound SoundFx("fx_bumper2",DOFContactors)*VolBump, 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
        Case 2 : PlaySound SoundFx("fx_bumper2",DOFContactors)*VolBump, 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
        Case 3 : PlaySound SoundFx("fx_bumper3",DOFContactors)*VolBump, 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
        Case 4 : PlaySound SoundFx("fx_bumper4",DOFContactors)*VolBump, 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    End Select
End Sub

Sub Woods_Hit (idx)
    PlaySound "woodhit", 0, Vol(ActiveBall)*VolWood, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub


 '************************************
 '          LEDs Display
 'Based on Scapino's 7 digit Reel LEDs
 '************************************


 LampTimer.Interval = 35
 LampTimer.Enabled = 1

 Sub LampTimer_Timer()
     UpdateLeds
     UpdateTextBoxes
 End Sub


 Dim Digits(32)
 Dim Patterns(11)
 Dim Patterns2(11)

 Patterns(0) = 0     'empty
 Patterns(1) = 63    '0
 Patterns(2) = 6     '1
 Patterns(3) = 91    '2
 Patterns(4) = 79    '3
 Patterns(5) = 102   '4
 Patterns(6) = 109   '5
 Patterns(7) = 125   '6
 Patterns(8) = 7     '7
 Patterns(9) = 127   '8
 Patterns(10) = 111  '9

 Patterns2(0) = 128  'empty
 Patterns2(1) = 191  '0
 Patterns2(2) = 134  '1
 Patterns2(3) = 219  '2
 Patterns2(4) = 207  '3
 Patterns2(5) = 230  '4
 Patterns2(6) = 237  '5
 Patterns2(7) = 253  '6
 Patterns2(8) = 135  '7
 Patterns2(9) = 255  '8
 Patterns2(10) = 239 '9

 Set Digits(0) = a0
 Set Digits(1) = a1
 Set Digits(2) = a2
 Set Digits(3) = a3
 Set Digits(4) = a4
 Set Digits(5) = a5
' Set Digits(6) = a6

 Set Digits(6) = b0
 Set Digits(7) = b1
 Set Digits(8) = b2
 Set Digits(9) = b3
 Set Digits(10) = b4
 Set Digits(11) = b5
' Set Digits(13) = b6

 Set Digits(12) = c0
 Set Digits(13) = c1
 Set Digits(14) = c2
 Set Digits(15) = c3
 Set Digits(16) = c4
 Set Digits(17) = c5
' Set Digits(20) = c6

 Set Digits(18) = d0
 Set Digits(19) = d1
 Set Digits(20) = d2
 Set Digits(21) = d3
 Set Digits(22) = d4
 Set Digits(23) = d5
' Set Digits(27) = d6

 Set Digits(24) = e0
 Set Digits(25) = e1
 Set Digits(26) = e2
 Set Digits(27) = e3

 Sub UpdateLeds
     On Error Resume Next
     Dim ChgLED, ii, jj, chg, stat
     ChgLED = Controller.ChangedLEDs(&HFF, &HFFFF)
     If Not IsEmpty(ChgLED)Then
         For ii = 0 To UBound(ChgLED)
             chg = chgLED(ii, 1):stat = chgLED(ii, 2)

             For jj = 0 to 10
                 If stat = Patterns(jj)OR stat = Patterns2(jj)then Digits(chgLED(ii, 0)).SetValue jj
             Next
         Next
     End IF
 End Sub


 Sub UpdateTextBoxes()
     NFadeT 13, l13, "BALL IN PLAY"
     NFadeT 27, l27, "MATCH"
     NFadeT 29, l29, "HIGH SCORE TO DATE"
     NFadeT 12, l12, "SAME PLAYER SHOOTS AGAIN" 'or 43
     NFadeT 45, l45, "GAME OVER"
     NFadeT 61, l61, "TILT"
     NFadeT 14, l14, "1"
     NFadeT 30, l30, "2"
     NFadeT 46, l46, "3"
     NFadeT 62, l62, "4"
End Sub

 Sub NFadeT(nr, a, b)
     Select Case controller.lamp(nr)
         Case False:a.Text = ""
         Case True:a.Text = b
     End Select
End Sub

CREDIT.Text = "CREDIT"


dim zz
If Table1.ShowDT = false then
    For each zz in DT:zz.Visible = false: Next
else
    For each zz in DT:zz.Visible = true: Next
End If


Sub Table1_Exit()

End Sub

' *******************************************************************************************************
' Positional Sound Playback Functions by DJRobX
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

Sub RollingSoundUpdate()
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



' Thalamus : Exit in a clean and proper way
Sub Table1_exit()
  Controller.Pause = False
  Controller.Stop
End Sub


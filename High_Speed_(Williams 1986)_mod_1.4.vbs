Option Explicit
Randomize

' Thalamus - Aug 2020 - improved SSF

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Const BallMass = 1.7
Const BallSize = 50

Const cGameName="hs_l4",UseSolenoids=2,UseLamps=0,UseGI=0,SSolenoidOn="SolOn",SSolenoidOff="SolOff", SCoin="coin"

LoadVPM "01500000", "S11.VBS", 3.10
Dim DesktopMode: DesktopMode = Table1.ShowDT

If DesktopMode = True Then 'Show Desktop components
SideRails.visible=1
LockdownBar.visible=1
'Ramp15.visible=1
Primitive13.visible=1
Else
SideRails.visible=0
LockdownBar.visible=0
'Ramp15.visible=0
Primitive13.visible=1
End if

'*************************************************************
'Solenoid Call backs
'**********************************************************************************************************
SolCallback(1)      = "bsTrough.SolIn"
SolCallback(2)      = "bstrough.SolOut"
SolCallback(3)      = "bsSaucer.SolOut"
SolCallback(5)      = "SetLamp 105," 'PF Light
SolCallback(6)      = "SetLamp 106," 'PF Light
SolCallback(7)      = "bsLeftLock.SolOut"           ' Left Hideout Eject
SolCallback(8)      = "bsRightLock.SolOut"          ' Right Hideout Eject
SolCallBack(9)      = "SetLamp 109," 'X2 Left Dome Flasher
SolCallback(11)     = "PFGI" 'General Illumination Relay
SolCallBack(12)     = "SetLamp 112," 'X2 Right Dome Flasher
SolCallback(13)     = "Divert"
SolCallback(14)     = "SolKickback"
SolCallback(15)=  "vpmSolSound SoundFX(""Knocker"",DOFKnocker),"
Solcallback(22)     = "SetLamp 122, " 'X2 Top flasher
SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

'///////////////////////-----General Sound Options-----///////////////////////
'//  VolumeDial:
'//  VolumeDial is the actual global volume multiplier for the mechanical sounds.
'//  Values smaller than 1 will decrease mechanical sounds volume.
'//  Recommended values should be no greater than 1.
Const VolumeDial = 0.8


Const PositionalSoundPlaybackConfiguration = 3
Dim FlipperUpAttackMinimumSoundLevel
Dim FlipperUpAttackMaximumSoundLevel
Dim FlipperUpSoundLevel
Dim FlipperDownSoundLevel
Dim FlipperLeftHitParm
Dim FlipperRightHitParm
Dim FlipperBallGuideSoundFactor
Dim SpinnerSoundLevel
Dim RubberStrongSoundFactor
Dim RubberWeakSoundFactor
Dim RubberFlipperSoundFactor
Dim WallImpactSoundFactor
Dim RollingSoundFactor

FlipperLeftHitParm = FlipperUpSoundLevel								'sound helper; not configurable
FlipperRightHitParm = FlipperUpSoundLevel								'sound helper; not configurable

FlipperBallGuideSoundFactor = 40										'volume multiplier; must not be zero

'/////////////////////////////  FLIPPER BALL GUIDE  ////////////////////////////
Sub RandomSoundFlipperBallGuide()
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 16 then
 		Select Case Int(Rnd*2)+1
			Case 1 : PlaySoundAtLevelActiveBall ("TOM_Apron_Hard_1"),  Vol(ActiveBall) * FlipperBallGuideSoundFactor
			Case 2 : PlaySoundAtLevelActiveBall ("TOM_Apron_Hard_2"),  Vol(ActiveBall) * 0.8 * FlipperBallGuideSoundFactor
		End Select
	End if
	If finalspeed >= 6 AND finalspeed <= 16 then
 		Select Case Int(Rnd*3)+1
			Case 1 : PlaySoundAtLevelActiveBall ("TOM_Apron_Medium_1"),  Vol(ActiveBall)
			Case 2 : PlaySoundAtLevelActiveBall ("TOM_Apron_Medium_2"),  Vol(ActiveBall) * FlipperBallGuideSoundFactor
			Case 3 : PlaySoundAtLevelActiveBall ("TOM_Apron_Medium_3"),  Vol(ActiveBall) * FlipperBallGuideSoundFactor
		End Select
 	End If
	If finalspeed < 6 Then
 		Select Case Int(Rnd*7)+1
			Case 1 : PlaySoundAtLevelActiveBall ("TOM_Apron_Soft_1"),  Vol(ActiveBall) * FlipperBallGuideSoundFactor
			Case 2 : PlaySoundAtLevelActiveBall ("TOM_Apron_Soft_2"),  Vol(ActiveBall) * FlipperBallGuideSoundFactor
			Case 3 : PlaySoundAtLevelActiveBall ("TOM_Apron_Soft_3"),  Vol(ActiveBall) * FlipperBallGuideSoundFactor
			Case 4 : PlaySoundAtLevelActiveBall ("TOM_Apron_Soft_4"),  Vol(ActiveBall) * FlipperBallGuideSoundFactor
			Case 5 : PlaySoundAtLevelActiveBall ("TOM_Apron_Soft_5"),  Vol(ActiveBall) * FlipperBallGuideSoundFactor
			Case 6 : PlaySoundAtLevelActiveBall ("TOM_Apron_Soft_6"),  Vol(ActiveBall) * FlipperBallGuideSoundFactor
			Case 7 : PlaySoundAtLevelActiveBall ("TOM_Apron_Soft_7"),  Vol(ActiveBall) * FlipperBallGuideSoundFactor
		End Select
	End if
End Sub

Sub SolLFlipper(Enabled)
     If Enabled Then
         LeftFlipper.RotateToEnd
		 RandomSoundFlipperUpLeft()
     Else
         RandomSoundFlipperDownLeft()
         FlipperLeftHitParm = FlipperUpSoundLevel
		 LeftFlipper.RotateToStart

         LeftFlipper.TimerEnabled = 1
         LeftFlipper.TimerInterval = 16
         LeftFlipper.return = returnspeed * 0.5

     End If
  End Sub

Sub SolRFlipper(Enabled)
     If Enabled Then
        RandomSoundFlipperUpRight()
         RightFlipper.RotateToEnd:RightFlipper1.RotateToEnd
     Else
         RandomSoundFlipperDownRight()
FlipperRightHitParm = FlipperUpSoundLevel
		 RightFlipper.RotateToStart:RightFlipper1.RotateToStart

         rightflipper.TimerEnabled = 1
         rightflipper.TimerInterval = 16
         rightflipper.return = returnspeed * 0.5

	 End If
End Sub


'*********** NFOZZY'S FLIPPERS *********************************
dim returnspeed, lfstep, rfstep
returnspeed = leftflipper.return
lfstep = 1
rfstep = 1

sub leftflipper_timer()
    select case lfstep
        Case 1: leftflipper.return = returnspeed * 0.6 :lfstep = lfstep + 1
        Case 2: leftflipper.return = returnspeed * 0.7 :lfstep = lfstep + 1
        Case 3: leftflipper.return = returnspeed * 0.8 :lfstep = lfstep + 1
        Case 4: leftflipper.return = returnspeed * 0.9 :lfstep = lfstep + 1
        Case 5: leftflipper.return = returnspeed * 1 :lfstep = lfstep + 1
        Case 6: leftflipper.timerenabled = 0 : lfstep = 1
    end select
end sub

sub rightflipper_timer()
    select case rfstep
        Case 1: rightflipper.return = returnspeed * 0.6 :rfstep = rfstep + 1
        Case 2: rightflipper.return = returnspeed * 0.7 :rfstep = rfstep + 1
        Case 3: rightflipper.return = returnspeed * 0.8 :rfstep = rfstep + 1
        Case 4: rightflipper.return = returnspeed * 0.9 :rfstep = rfstep + 1
        Case 5: rightflipper.return = returnspeed * 1 :rfstep = rfstep + 1
        Case 6: rightflipper.timerenabled = 0 : rfstep = 1
    end select
end sub

'**********************************************************************************************************
 'Solenoid Controlled toys
'**********************************************************************************************************

Sub SolKickBack(enabled)
    If enabled Then
       Plunger1.Fire
       PlaySound SoundFX("Popper",DOFContactors)
    Else
       Plunger1.PullBack
    End If
End Sub

Sub Divert(enabled)
    If Enabled Then
        Diverter1.IsDropped = 0
        Diverter2.IsDropped = 0
        PrimFlipper1.roty = -90
        PrimFlipper2.roty = -90
        PlaySound SoundFX("sc_loop2 2",DOFContactors)
    Else
        Diverter1.IsDropped = 1
        Diverter2.IsDropped = 1
        PrimFlipper1.roty = 0
        PrimFlipper2.roty = 0
        PlaySound SoundFX("sc_loop2 2",DOFContactors)
    End If
End Sub

'Playfield GI
Sub PFGI(Enabled)
    If Enabled Then
        dim xx
        For each xx in GI:xx.State = 0: Next
        PlaySound "fx_relay"
    Else
        For each xx in GI:xx.State = 1: Next
        PlaySound "fx_relay"
    End If
End Sub

'**********************************************************************************************************
 'Initiate Table
'**********************************************************************************************************

Dim bsTrough, bsSaucer, bsLeftLock, bsRightLock, SubSpeed
 ' using table width and height in script slows down the performance
dim tablewidth: tablewidth = Table1.width
dim tableheight: tableheight = Table1.height
Sub Table1_Init
    vpmInit Me
    On Error Resume Next
        With Controller
        .GameName = cGameName
        If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
        .SplashInfoLine = "High Speed (Williams)"&chr(13)&"You Suck"
        .HandleMechanics=0
        .HandleKeyboard=0
        .ShowDMDOnly=1
        .ShowFrame=0
        .ShowTitle=0
        .hidden = 1
         On Error Resume Next
         Controller.SolMask(0)=0
         vpmTimer.AddTimer 4000,"Controller.SolMask(0)=&Hffffffff'"
         .Run GetPlayerHWnd
         If Err Then MsgBox Err.Description
         On Error Goto 0
     End With
     On Error Goto 0

    PinMAMETimer.Interval=PinMAMEInterval
    PinMAMETimer.Enabled=1

    vpmNudge.TiltSwitch=1
    vpmNudge.Sensitivity=3
    vpmNudge.TiltObj=Array(Bumper1,Bumper2,Bumper3,LeftSlingshot,RightSlingshot)

    Set bsTrough=New cvpmBallStack
        bsTrough.InitSw 9,12,11,10,0,0,0,0
        bsTrough.InitKick BallRelease,90,10
        bsTrough.InitExitSnd SoundFX("ballrelease",DOFContactors), SoundFX("Solenoid",DOFContactors)
        bsTrough.Balls=3

    Set bsSaucer = New cvpmBallStack
        bsSaucer.InitSaucer sw16,16,96,5
        bsSaucer.InitExitSnd SoundFX("Popper",DOFContactors), SoundFX("Solenoid",DOFContactors)

    Set bsLeftLock = New cvpmBallStack
        bsLeftLock.InitSw 0,0,40,0,0,0,0,0
        bsLeftLock.InitSaucer LKick,40, 0,25
		bsLeftLock.KickforceVar = 5
        bsLeftLock.InitExitSnd SoundFX("kicker2",DOFContactors), SoundFX("Solenoid",DOFContactors)

     Set bsRightLock = New cvpmBallStack
        bsRightLock.InitSw 0,0,48,0,0,0,0,0
        bsRightLock.InitSaucer RKick,48,0,25
		bsRightLock.KickforceVar = 5
        bsRightLock.InitExitSnd SoundFX("kicker2",DOFContactors), SoundFX("Solenoid",DOFContactors)

    Plunger1.Pullback
    Diverter1.IsDropped = 1
    Diverter2.IsDropped = 1

 End Sub

'**********************************************************************************************************
'Plunger code
'**********************************************************************************************************
Sub Table1_KeyDown(ByVal keycode)
    If keycode = LeftTiltKey Then Nudge 90, 5:PlaySoundAt SoundFX("fx_nudge",0), CardLeft
    If keycode = RightTiltKey Then Nudge 270, 5:PlaySoundAt SoundFX("fx_nudge",0), CardRight
    If keycode = CenterTiltKey Then Nudge 0, 6:PlaySoundAt SoundFX("fx_nudge",0), Drain
    If vpmKeyDown(keycode) Then Exit Sub
End Sub

Sub Table1_KeyDown(ByVal KeyCode)
    If KeyDownHandler(keycode) Then Exit Sub
    If keycode = PlungerKey Then Plunger.Pullback:playsoundAtVol "plungerpull", Plunger, VolumeDial
End Sub

Sub Table1_KeyUp(ByVal KeyCode)
    If KeyUpHandler(keycode) Then Exit Sub
    If keycode = PlungerKey Then Plunger.Fire:PlaySoundAtVol "plunger", Plunger, VolumeDial
End Sub

'**********************************************************************************************************

 ' Drain hole and kickers
Sub Drain_Hit:bsTrough.addball me : playsound"drain" : End Sub
Sub sw16_Hit:bsSaucer.addball 0 : playsound"popper_ball" : End Sub
Sub LKick_Hit:bsLeftLock.AddBall 0 : playsound"Target" : End Sub
Sub RKick_Hit:bsRightLock.AddBall 0 : playsound"Target" : End Sub

'fake 180 turn wire Ramp
Sub kicker1_Hit
    SubSpeed=ABS(ActiveBall.VelY)
    kicker1.DestroyBall
    kicker2.CreateSizedballWithMass Ballsize/2,BallMass
    kicker2.Kick 180,SQR(SubSpeed)
End Sub

Sub kicker3_Hit
    SubSpeed=ABS(ActiveBall.VelY)
    kicker3.DestroyBall
    kicker4.CreateSizedballWithMass Ballsize/2,BallMass
    kicker4.Kick 180,SQR(SubSpeed)
End Sub

'Plastic Triggers
Sub Trigger111_Hit():PlaySoundAtVol "fx_lr7", ActiveBall, 1: End Sub
Sub Trigger112_Hit():PlaySoundAtVol "fx_lr3", ActiveBall, 1: End Sub
Sub Trigger113_Hit():PlaySoundAtVol "fx_lr4", ActiveBall, 1: End Sub
Sub Trigger114_Hit():PlaySoundAtVol "fx_lr2", ActiveBall, 1: End Sub

RubberStrongSoundFactor = 0.8							'volume multiplier; must not be zero
RubberWeakSoundFactor = 0.9					'volume multiplier; must not be zero
RubberFlipperSoundFactor = 0.2


'/////////////////////////////  RUBBERS AND POSTS - STRONG IMPACTS  ////////////////////////////
Sub RandomSoundRubberStrong()
	Select Case Int(Rnd*10)+1
		Case 1 : PlaySoundAtLevelActiveBall ("TOM_Rubber_Flipper_Slingshot_Strong_1"), Vol(ActiveBall) * RubberStrongSoundFactor
		Case 2 : PlaySoundAtLevelActiveBall ("TOM_Rubber_Flipper_Slingshot_Strong_2"), Vol(ActiveBall) * RubberStrongSoundFactor
		Case 3 : PlaySoundAtLevelActiveBall ("TOM_Rubber_Flipper_Slingshot_Strong_3"), Vol(ActiveBall) * RubberStrongSoundFactor
		Case 4 : PlaySoundAtLevelActiveBall ("TOM_Rubber_Flipper_Slingshot_Strong_4"), Vol(ActiveBall) * RubberStrongSoundFactor
		Case 5 : PlaySoundAtLevelActiveBall ("TOM_Rubber_Flipper_Slingshot_Strong_5"), Vol(ActiveBall) * RubberStrongSoundFactor
		Case 6 : PlaySoundAtLevelActiveBall ("TOM_Rubber_Flipper_Slingshot_Strong_6"), Vol(ActiveBall) * RubberStrongSoundFactor
		Case 7 : PlaySoundAtLevelActiveBall ("TOM_Rubber_Flipper_Slingshot_Strong_7"), Vol(ActiveBall) * RubberStrongSoundFactor
		Case 8 : PlaySoundAtLevelActiveBall ("TOM_Rubber_Flipper_Slingshot_Strong_8"), Vol(ActiveBall) * RubberStrongSoundFactor
		Case 9 : PlaySoundAtLevelActiveBall ("TOM_Rubber_Flipper_Slingshot_Strong_9"), Vol(ActiveBall) * RubberStrongSoundFactor
		Case 10 : PlaySoundAtLevelActiveBall ("TOM_Rubber_1_Hard"), Vol(ActiveBall) * RubberStrongSoundFactor * 0.6
	End Select
End Sub

'/////////////////////////////  RUBBERS AND POSTS - WEAK IMPACTS  ////////////////////////////
Sub RandomSoundRubberWeak()
	Select Case Int(Rnd*9)+1
		Case 1 : PlaySoundAtLevelActiveBall ("TOM_Rubber_1"), Vol(ActiveBall) * RubberWeakSoundFactor
		Case 2 : PlaySoundAtLevelActiveBall ("TOM_Rubber_2"), Vol(ActiveBall) * RubberWeakSoundFactor
		Case 3 : PlaySoundAtLevelActiveBall ("TOM_Rubber_3"), Vol(ActiveBall) * RubberWeakSoundFactor
		Case 4 : PlaySoundAtLevelActiveBall ("TOM_Rubber_3"), Vol(ActiveBall) * RubberWeakSoundFactor
		Case 5 : PlaySoundAtLevelActiveBall ("TOM_Rubber_5"), Vol(ActiveBall) * RubberWeakSoundFactor
		Case 6 : PlaySoundAtLevelActiveBall ("TOM_Rubber_6"), Vol(ActiveBall) * RubberWeakSoundFactor
		Case 7 : PlaySoundAtLevelActiveBall ("TOM_Rubber_7"), Vol(ActiveBall) * RubberWeakSoundFactor
		Case 8 : PlaySoundAtLevelActiveBall ("TOM_Rubber_8"), Vol(ActiveBall) * RubberWeakSoundFactor
		Case 9 : PlaySoundAtLevelActiveBall ("TOM_Rubber_9"), Vol(ActiveBall) * RubberWeakSoundFactor
	End Select
End Sub

Dim MetalImpactSoundFactor
MetalImpactSoundFactor = 0.2


'/////////////////////////////  METAL TOUCH SOUNDS  ////////////////////////////
Sub RandomSoundMetal()
	Select Case Int(Rnd*13)+1
		Case 1 : PlaySoundAtLevelActiveBall ("TOM_Metal_Touch_1"), Vol(ActiveBall) * MetalImpactSoundFactor
		Case 2 : PlaySoundAtLevelActiveBall ("TOM_Metal_Touch_2"), Vol(ActiveBall) * MetalImpactSoundFactor
		Case 3 : PlaySoundAtLevelActiveBall ("TOM_Metal_Touch_3"), Vol(ActiveBall) * MetalImpactSoundFactor
		Case 4 : PlaySoundAtLevelActiveBall ("TOM_Metal_Touch_4"), Vol(ActiveBall) * MetalImpactSoundFactor
		Case 5 : PlaySoundAtLevelActiveBall ("TOM_Metal_Touch_5"), Vol(ActiveBall) * MetalImpactSoundFactor
		Case 6 : PlaySoundAtLevelActiveBall ("TOM_Metal_Touch_6"), Vol(ActiveBall) * MetalImpactSoundFactor
		Case 7 : PlaySoundAtLevelActiveBall ("TOM_Metal_Touch_7"), Vol(ActiveBall) * MetalImpactSoundFactor
		Case 8 : PlaySoundAtLevelActiveBall ("TOM_Metal_Touch_8"), Vol(ActiveBall) * MetalImpactSoundFactor
		Case 9 : PlaySoundAtLevelActiveBall ("TOM_Metal_Touch_9"), Vol(ActiveBall) * MetalImpactSoundFactor
		Case 10 : PlaySoundAtLevelActiveBall ("TOM_Metal_Touch_10"), Vol(ActiveBall) * MetalImpactSoundFactor
		Case 11 : PlaySoundAtLevelActiveBall ("TOM_Metal_Touch_11"), Vol(ActiveBall) * MetalImpactSoundFactor
		Case 12 : PlaySoundAtLevelActiveBall ("TOM_Metal_Touch_12"), Vol(ActiveBall) * MetalImpactSoundFactor
		Case 13 : PlaySoundAtLevelActiveBall ("TOM_Metal_Touch_13"), Vol(ActiveBall) * MetalImpactSoundFactor
	End Select
End Sub


'/////////////////////////////  WALL IMPACTS EVENTS  ////////////////////////////
Sub Wall21_Hit()
	RandomSoundMetal()
End Sub

Sub Wall26_Hit()
	RandomSoundMetal()
End Sub

Sub Wall449_Hit()
	RandomSoundMetal()
End Sub

Sub Wall375_Hit()
	RandomSoundMetal()
End Sub

Sub Wall374_Hit()
	RandomSoundMetal()
End Sub

Sub Wall3_Hit()
	RandomSoundMetal()
End Sub

Sub Wall6_Hit()
	RandomSoundMetal()
End Sub

Sub Wall519_Hit()
	RandomSoundMetal()
End Sub

Sub Wall4_Hit()
	RandomSoundFlipperBallGuide()
End Sub

Sub Wall28_Hit()
	RandomSoundFlipperBallGuide()
End Sub

Sub Wall36_Hit()
	RandomSoundMetal()
End Sub

Sub Wall6_Hit()
	RandomSoundMetal()
End Sub

'Sub Wall454_Hit()
'	RandomSoundFlipperBallGuide()
'End Sub

Sub Wall520_Hit()
	RandomSoundMetal()
End Sub

Sub Wall42_Hit()
	RandomSoundMetal()
End Sub

Sub Wall12_Hit()
	RandomSoundFlipperBallGuide()
End Sub


'Wire Triggers
Sub SW20_Hit : Controller.Switch(20)=1 : playsoundAtVol"rollover" , ActiveBall, 1: End Sub
Sub SW20_unHit : Controller.Switch(20)=0:End Sub
Sub SW21_Hit : Controller.Switch(21)=1 : playsoundAtVol"rollover" , ActiveBall, 1: End Sub
Sub SW21_unHit : Controller.Switch(21)=0:End Sub
Sub SW31_Hit : Controller.Switch(31)=1 : playsoundAtVol"rollover" , ActiveBall, 1: End Sub
Sub SW31_unHit : Controller.Switch(31)=0 : End Sub
Sub SW32_Hit : Controller.Switch(32)=1 : playsoundAtVol"rollover" , ActiveBall, 1: End Sub
Sub SW32_unHit : Controller.Switch(32)=0:End Sub
Sub SW36_Hit : Controller.Switch(36)=1 : playsoundAtVol"rollover" , ActiveBall, 1: End Sub
Sub SW36_unHit : Controller.Switch(36)=0:End Sub
sub Sw39_hit:controller.Switch(39) = 1 : playsoundAtVol"rollover" , ActiveBall, 1: End Sub
sub Sw39_unhit:controller.Switch(39) = 0 : end sub
sub Sw47_hit:controller.Switch(47) = 1 : playsoundAtVol"rollover" , ActiveBall, 1: End Sub
sub Sw47_unhit:controller.Switch(47) = 0 : end sub

'Stand Up Targets

TargetSoundFactor = 0.005							'volume multiplier; must not be zero
BallBouncePlayfieldSoftFactor = 0.005			'volume multiplier; must not be zero

Sub sw17_Hit
	vpmTimer.PulseSw 17
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 10 then
 		RandomSoundTargetHitStrong()
		RandomSoundBallBouncePlayfieldSoft()
	End if
	If finalspeed <= 10 then
 		RandomSoundTargetHitWeak()
 	End If
End Sub

Sub sw18_Hit
	vpmTimer.PulseSw 18
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 10 then
 		RandomSoundTargetHitStrong()
		RandomSoundBallBouncePlayfieldSoft()
	End if
	If finalspeed <= 10 then
 		RandomSoundTargetHitWeak()
 	End If
End Sub

Sub sw19_Hit
	vpmTimer.PulseSw 19
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 10 then
 		RandomSoundTargetHitStrong()
		RandomSoundBallBouncePlayfieldSoft()
	End if
	If finalspeed <= 10 then
 		RandomSoundTargetHitWeak()
 	End If
End Sub

Sub sw13_Hit
	vpmTimer.PulseSw 13
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 10 then
 		RandomSoundTargetHitStrong()
		RandomSoundBallBouncePlayfieldSoft()
	End if
	If finalspeed <= 10 then
 		RandomSoundTargetHitWeak()
 	End If
End Sub

Sub sw14_Hit
	vpmTimer.PulseSw 14
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 10 then
 		RandomSoundTargetHitStrong()
		RandomSoundBallBouncePlayfieldSoft()
	End if
	If finalspeed <= 10 then
 		RandomSoundTargetHitWeak()
 	End If
End Sub

Sub sw15_Hit
	vpmTimer.PulseSw 15
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 10 then
 		RandomSoundTargetHitStrong()
		RandomSoundBallBouncePlayfieldSoft()
	End if
	If finalspeed <= 10 then
 		RandomSoundTargetHitWeak()
 	End If
End Sub

Sub sw22_Hit
	vpmTimer.PulseSw 22
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 10 then
 		RandomSoundTargetHitStrong()
		RandomSoundBallBouncePlayfieldSoft()
	End if
	If finalspeed <= 10 then
 		RandomSoundTargetHitWeak()
 	End If
End Sub

Sub sw23_Hit
	vpmTimer.PulseSw 23
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 10 then
 		RandomSoundTargetHitStrong()
		RandomSoundBallBouncePlayfieldSoft()
	End if
	If finalspeed <= 10 then
 		RandomSoundTargetHitWeak()
 	End If
End Sub

Sub sw24_Hit
	vpmTimer.PulseSw 24
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 10 then
 		RandomSoundTargetHitStrong()
		RandomSoundBallBouncePlayfieldSoft()
	End if
	If finalspeed <= 10 then
 		RandomSoundTargetHitWeak()
 	End If
End Sub

Dim BallBouncePlayfieldSoftFactor
Dim BallBouncePlayfieldHardFactor
Dim TargetSoundFactor

Sub PlaySoundAtLevelActiveBall(playsoundparams, aVol)
	PlaySound playsoundparams, 0, aVol * VolumeDial, AudioPan(ActiveBall), 0, 0, 0, 0, AudioFade(ActiveBall)
End Sub


'/////////////////////////////  TARGET HIT SOUNDS  ////////////////////////////
Sub RandomSoundTargetHitStrong()
	Select Case Int(Rnd*4)+1
		Case 1 : PlaySoundAtLevelActiveBall SoundFX("TOM_Target_Hit_5",DOFTargets), 100*Vol(ActiveBall) * 0.45 * TargetSoundFactor
		Case 2 : PlaySoundAtLevelActiveBall SoundFX("TOM_Target_Hit_6",DOFTargets), 100*Vol(ActiveBall) * 0.45 * TargetSoundFactor
		Case 3 : PlaySoundAtLevelActiveBall SoundFX("TOM_Target_Hit_7",DOFTargets), 100*Vol(ActiveBall) * 0.45 * TargetSoundFactor
		Case 4 : PlaySoundAtLevelActiveBall SoundFX("TOM_Target_Hit_8",DOFTargets), 100*Vol(ActiveBall) * 0.45 * TargetSoundFactor
	End Select
End Sub

Sub RandomSoundTargetHitWeak()
	Select Case Int(Rnd*4)+1
		Case 1 : PlaySoundAtLevelActiveBall SoundFX("TOM_Target_Hit_1",DOFTargets), 100*Vol(ActiveBall) * TargetSoundFactor
		Case 2 : PlaySoundAtLevelActiveBall SoundFX("TOM_Target_Hit_2",DOFTargets), 100*Vol(ActiveBall) * TargetSoundFactor
		Case 3 : PlaySoundAtLevelActiveBall SoundFX("TOM_Target_Hit_3",DOFTargets), 100*Vol(ActiveBall) * TargetSoundFactor
		Case 4 : PlaySoundAtLevelActiveBall SoundFX("TOM_Target_Hit_4",DOFTargets), 100*Vol(ActiveBall) * TargetSoundFactor
	End Select
End Sub

'/////////////////////////////  BALL BOUNCE SOUNDS  ////////////////////////////
Sub RandomSoundBallBouncePlayfieldSoft()
	Select Case Int(Rnd*9)+1
		Case 1 : PlaySoundAtLevelActiveBall ("TOM_Ball_Bounce_Playfield_Soft_1"), 100*Vol(ActiveBall) * BallBouncePlayfieldSoftFactor
		Case 2 : PlaySoundAtLevelActiveBall ("TOM_Ball_Bounce_Playfield_Soft_2"), 100*Vol(ActiveBall) * BallBouncePlayfieldSoftFactor * 0.5
		Case 3 : PlaySoundAtLevelActiveBall ("TOM_Ball_Bounce_Playfield_Soft_3"), 100*Vol(ActiveBall) * BallBouncePlayfieldSoftFactor * 0.8
		Case 4 : PlaySoundAtLevelActiveBall ("TOM_Ball_Bounce_Playfield_Soft_4"), 100*Vol(ActiveBall) * BallBouncePlayfieldSoftFactor * 0.5
		Case 5 : PlaySoundAtLevelActiveBall ("TOM_Ball_Bounce_Playfield_Soft_5"), 100*Vol(ActiveBall) * BallBouncePlayfieldSoftFactor
		Case 6 : PlaySoundAtLevelActiveBall ("TOM_Ball_Bounce_Playfield_Hard_1"), 100*Vol(ActiveBall) * BallBouncePlayfieldSoftFactor * 0.2
		Case 7 : PlaySoundAtLevelActiveBall ("TOM_Ball_Bounce_Playfield_Hard_2"), 100*Vol(ActiveBall) * BallBouncePlayfieldSoftFactor * 0.2
		Case 8 : PlaySoundAtLevelActiveBall ("TOM_Ball_Bounce_Playfield_Hard_5"), 100*Vol(ActiveBall) * BallBouncePlayfieldSoftFactor * 0.2
		Case 9 : PlaySoundAtLevelActiveBall ("TOM_Ball_Bounce_Playfield_Hard_7"), 100*Vol(ActiveBall) * BallBouncePlayfieldSoftFactor * 0.3
	End Select
End Sub

Sub RandomSoundBallBouncePlayfieldHard()
	Select Case Int(Rnd*7)+1
		Case 1 : PlaySoundAtLevelActiveBall ("TOM_Ball_Bounce_Playfield_Hard_1"), 100*Vol(ActiveBall) * BallBouncePlayfieldHardFactor
		Case 2 : PlaySoundAtLevelActiveBall ("TOM_Ball_Bounce_Playfield_Hard_2"), 100*Vol(ActiveBall) * BallBouncePlayfieldHardFactor
		Case 3 : PlaySoundAtLevelActiveBall ("TOM_Ball_Bounce_Playfield_Hard_3"),100* Vol(ActiveBall) * BallBouncePlayfieldHardFactor
		Case 4 : PlaySoundAtLevelActiveBall ("TOM_Ball_Bounce_Playfield_Hard_4"), 100*Vol(ActiveBall) * BallBouncePlayfieldHardFactor
		Case 5 : PlaySoundAtLevelActiveBall ("TOM_Ball_Bounce_Playfield_Hard_5"), 100*Vol(ActiveBall) * BallBouncePlayfieldHardFactor
		Case 6 : PlaySoundAtLevelActiveBall ("TOM_Ball_Bounce_Playfield_Hard_6"), 100*Vol(ActiveBall) * BallBouncePlayfieldHardFactor
		Case 7 : PlaySoundAtLevelActiveBall ("TOM_Ball_Bounce_Playfield_Hard_7"), 100*Vol(ActiveBall) * BallBouncePlayfieldHardFactor
	End Select
End Sub


Sub sw25_Hit
	vpmTimer.PulseSw 25
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 10 then
 		RandomSoundTargetHitStrong()
		RandomSoundBallBouncePlayfieldSoft()
	End if
	If finalspeed <= 10 then
 		RandomSoundTargetHitWeak()
 	End If
End Sub

Sub sw26_Hit
	vpmTimer.PulseSw 26
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 10 then
 		RandomSoundTargetHitStrong()
		RandomSoundBallBouncePlayfieldSoft()
	End if
	If finalspeed <= 10 then
 		RandomSoundTargetHitWeak()
 	End If
End Sub

Sub sw27_Hit
	vpmTimer.PulseSw 27
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 10 then
 		RandomSoundTargetHitStrong()
		RandomSoundBallBouncePlayfieldSoft()
	End if
	If finalspeed <= 10 then
 		RandomSoundTargetHitWeak()
 	End If
End Sub

Sub sw28_Hit
	vpmTimer.PulseSw 28
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 10 then
 		RandomSoundTargetHitStrong()
		RandomSoundBallBouncePlayfieldSoft()
	End if
	If finalspeed <= 10 then
 		RandomSoundTargetHitWeak()
 	End If
End Sub

 Sub sw29_Hit
	vpmTimer.PulseSw 29
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 10 then
 		RandomSoundTargetHitStrong()
		RandomSoundBallBouncePlayfieldSoft()
	End if
	If finalspeed <= 10 then
 		RandomSoundTargetHitWeak()
 	End If
End Sub

Sub sw30_Hit
	vpmTimer.PulseSw 30
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 10 then
 		RandomSoundTargetHitStrong()
		RandomSoundBallBouncePlayfieldSoft()
	End if
	If finalspeed <= 10 then
 		RandomSoundTargetHitWeak()
 	End If
End Sub


'/////////////////////////////  BUMPER SOLENOID SOUNDS  ////////////////////////////
Sub RandomSoundBumperA()
	Select Case Int(Rnd*5)+1
		Case 1 : PlaySoundAtLevelStatic SoundFX("TOM_Bumpers_Reworked_v2_Top_1",DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bumper1
		Case 2 : PlaySoundAtLevelStatic SoundFX("TOM_Bumpers_Reworked_v2_Top_2",DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bumper1
		Case 3 : PlaySoundAtLevelStatic SoundFX("TOM_Bumpers_Reworked_v2_Top_3",DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bumper1
		Case 4 : PlaySoundAtLevelStatic SoundFX("TOM_Bumpers_Reworked_v2_Top_4",DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bumper1
		Case 5 : PlaySoundAtLevelStatic SoundFX("TOM_Bumpers_Reworked_v2_Top_5",DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bumper1
	End Select
End Sub

Sub RandomSoundBumperB()
	Select Case Int(Rnd*5)+1
		Case 1 : PlaySoundAtLevelStatic SoundFX("TOM_Bumpers_Reworked_v2_Middle_1",DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bumper2
		Case 2 : PlaySoundAtLevelStatic SoundFX("TOM_Bumpers_Reworked_v2_Middle_2",DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bumper2
		Case 3 : PlaySoundAtLevelStatic SoundFX("TOM_Bumpers_Reworked_v2_Middle_3",DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bumper2
		Case 4 : PlaySoundAtLevelStatic SoundFX("TOM_Bumpers_Reworked_v2_Middle_4",DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bumper2
		Case 5 : PlaySoundAtLevelStatic SoundFX("TOM_Bumpers_Reworked_v2_Middle_5",DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bumper2
	End Select
End Sub

Sub RandomSoundBumperC()
	Select Case Int(Rnd*5)+1
		Case 1 : PlaySoundAtLevelStatic SoundFX("TOM_Bumpers_Reworked_v2_Bottom_1",DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bumper3
		Case 2 : PlaySoundAtLevelStatic SoundFX("TOM_Bumpers_Reworked_v2_Bottom_2",DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bumper3
		Case 3 : PlaySoundAtLevelStatic SoundFX("TOM_Bumpers_Reworked_v2_Bottom_3",DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bumper3
		Case 4 : PlaySoundAtLevelStatic SoundFX("TOM_Bumpers_Reworked_v2_Bottom_4",DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bumper3
		Case 5 : PlaySoundAtLevelStatic SoundFX("TOM_Bumpers_Reworked_v2_Bottom_5",DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bumper3
	End Select
End Sub


'******************************************************
'					BUMPERS
'******************************************************

Dim dirRing1 : dirRing1 = -1
Dim dirRing2 : dirRing2 = -1
Dim dirRing3 : dirRing3 = -1


Sub Bumper1_Hit
vpmTimer.PulseSw 33
RandomSoundBumperA()
End Sub

Sub Bumper2_Hit
vpmTimer.PulseSw 34
RandomSoundBumperB()
End Sub

Sub Bumper3_Hit
vpmTimer.PulseSw 35
RandomSoundBumperC()
End Sub

Sub Bumper1_timer()
	BR1.Z = BR1.Z + (5 * dirRing1)
	BR4.Y = mdistY - SQR((ABS(BR1.Y)-mdistY)^2 +(BR1.Z)^2) * dCos(mangle) + BR1.Z * dSin(mangle)
	BR4.Z = BR1.Z * dCos(mangle) + SQR((ABS(BR1.Y)-mdistY)^2 +(BR1.Z)^2) * dSin(mangle)
	If BR1.Z <= 0 Then dirRing1 = 1
	If BR1.Z >= 40 Then
		dirRing1 = -1
		BR1.Z = 40
		Me.TimerEnabled = 0
	End If
End Sub

Sub Bumper2_timer()
	BR2.Z = BR2.Z + (5 * dirRing2)
	If BR2.Z <= 0 Then dirRing2 = 1
	If BR2.Z >= 40 Then
		dirRing2 = -1
		BR2.Z = 40
		Me.TimerEnabled = 0
	End If
End Sub

Sub Bumper3_timer()
	BR3.Z = BR3.Z + (5 * dirRing3)
	If BR3.Z <= 0 Then dirRing3 = 1
	If BR3.Z >= 40 Then
		dirRing3 = -1
		BR3.Z = 40
		Me.TimerEnabled = 0
	End If
End Sub

'Ramp Triggers
Sub SW42_Hit:Controller.Switch(42)=1:End Sub
Sub SW42_unHit:Controller.Switch(42)=0:End Sub
Sub SW43_Hit:Controller.Switch(43)=1:End Sub
Sub SW43_unHit:Controller.Switch(43)=0:End Sub

'
' '/////////////////////////////  SPINNER  ////////////////////////////
' Sub SoundSpinner()
' 	PlaySoundAtLevelStatic ("TOM_NK_Spinner_12"), SpinnerSoundLevel, sw44
' End Sub
'
' '/////////////////////////////  SPINNER  ////////////////////////////
' Sub SoundSpinner()
' 	PlaySoundAtLevelStatic ("TOM_NK_Spinner_12"), SpinnerSoundLevel, sw45
' End Sub
'
' '/////////////////////////////  SPINNER  ////////////////////////////
' Sub SoundSpinner()
' 	PlaySoundAtLevelStatic ("TOM_NK_Spinner_12"), SpinnerSoundLevel, sw46
' End Sub
'
'
' SpinnerSoundLevel = 0.01											'volume level; range [0, 1]


'Spinner
Sub sw44_Spin:vpmTimer.PulseSw 44 : PlaySoundAtVol "TOM_NK_Spinner_12", sw44, 0.01 : End Sub
Sub sw45_Spin:vpmTimer.PulseSw 45 : PlaySoundAtVol "TOM_NK_Spinner_12", sw45, 0.01 : End Sub
Sub sw46_Spin:vpmTimer.PulseSw 46 : PlaySoundAtVol "TOM_NK_Spinner_12", sw46, 0.01 : End Sub

'Star Triggers
Sub SW51_Hit:Controller.Switch(51)=1 : playsoundAtVol"rollover" , ActiveBall, 1: End Sub
Sub SW51_unHit:Controller.Switch(51)=0:End Sub
Sub SW52_Hit:Controller.Switch(52)=1 : playsoundAtVol"rollover" , ActiveBall, 1: End Sub
Sub SW52_unHit:Controller.Switch(52)=0:End Sub

'Generic Ramp Sounds
Sub Trigger1_Hit : playsoundAtVol"Wire Ramp" , ActiveBall, 1: End Sub
Sub Trigger2_Hit : playsoundAtVol"Wire Ramp" , ActiveBall, 1: End Sub
Sub Trigger3_Hit : playsoundAtVol"vuk_exit" , ActiveBall, 1: End Sub
Sub Trigger4_Hit : playsoundAtVol"vuk_exit" , ActiveBall, 1: End Sub
Sub Trigger5_Hit : stopSound "Wire Ramp"
                   playsoundAtVol"WireRamp_Hit1", ActiveBall, 1: End Sub
Sub Trigger6_Hit : stopSound "Wire Ramp"
                   playsoundAtVol"WireRamp_Hit2",ActiveBall,1: End Sub
Sub Trigger7_Hit : playsoundAtVol"Ball Drop" , ActiveBall, 1: End Sub
Sub Trigger8_Hit : playsoundAtVol"Ball Drop" , ActiveBall, 1: End Sub
Sub Trigger9_Hit : playsoundAtVol"Ball Drop" , ActiveBall, 1: End Sub
Sub Trigger10_Hit : playsoundAtVol"WireRamp_Hit" , ActiveBall, 1: End Sub
Sub Trigger11_Hit : playsoundAtVol"WireRamp_Hit" , ActiveBall, 1: End Sub
Sub Trigger12_Hit : playsoundAtVol"WireRamp_Hit1" , ActiveBall, 1: End Sub
Sub Trigger13_Hit : playsoundAtVol"WireRamp_Hit1" , ActiveBall, 1: End Sub


'//  RelaysPosition:
'//  1 = Relays positioned with power board (Provides sound spread from the left and right back channels)
'//  2 = Relays positioned with GI strings (Provides sound spread from left/right/front/rear surround channels)
Const RelaysPosition = 2


'/////////////////////////////  SUPPORTING BALL & SOUND FUNCTIONS  ////////////////////////////
Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "table1" is the name of the table
	Dim tmp
	Select Case PositionalSoundPlaybackConfiguration
		Case 1
			AudioFade = 0
		Case 2
			AudioFade = 0
		Case 3
			tmp = tableobj.y * 2 / tableheight-1
			'tmp = tableobj.x * 2 / table1.height-1
			If tmp > 0 Then
				AudioFade = Csng(tmp ^10)
			Else
				AudioFade = Csng(-((- tmp) ^10) )
			End If
	End Select
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "table1" is the name of the table
	Dim tmp
	Select Case PositionalSoundPlaybackConfiguration
		Case 1
			AudioPan = 0
		Case 2
			tmp = tableobj.x * 2 / tablewidth-1
		    'tmp = tableobj.x * 2 / table1.width-1
			If tmp > 0 Then
				AudioPan = Csng(tmp ^10)
			Else
				AudioPan = Csng(-((- tmp) ^10) )
			End If
		Case 3
			tmp = tableobj.x * 2 / tablewidth-1
			'tmp = tableobj.x * 2 / table1.width-1
			If tmp > 0 Then
				AudioPan = Csng(tmp ^10)
			Else
				AudioPan = Csng(-((- tmp) ^10) )
			End If
	End Select
End Function

Function Vol(ball) ' Calculates the volume of the sound based on the ball speed
	Vol = Csng(BallVel(ball) ^2)
	'Vol = Csng(BallVel(ball) ^2/55)
	'Vol = Csng(BallVel(ball) ^2)
End Function

Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
    Pitch = BallVel(ball) * 20
End Function

Function BallVel(ball) 'Calculates the ball speed
    BallVel = INT(SQR((ball.VelX ^2) + (ball.VelY ^2) ) )
End Function

Function VolPlayfieldRoll(ball) ' Calculates the roll volume of the sound based on the ball speed
	VolPlayfieldRoll = RollingSoundFactor * 0.0005 * Csng(BallVel(ball) ^3)
End Function

Function VolPlasticRampRoll(ball) ' Calculates the roll volume of the sound based on the ball speed
	VolPlasticRampRoll = RollingSoundFactor * 0.0005 * Csng(BallVel(ball) ^3)
End Function

Function PitchPlayfieldRoll(ball) ' Calculates the roll pitch of the sound based on the ball speed
    PitchPlayfieldRoll = BallVel(ball) ^2 * 15
End Function

Function PitchPlasticRamp(ball) ' Calculates the pitch of the sound based on the ball speed - used for plastic ramps roll sound
    PitchPlasticRamp = BallVel(ball) * 20
End Function

'/////////////////////////////  SOUND PLAYBACK FUNCTIONS  ////////////////////////////
'/////////////////////////////  POSITIONAL SOUND PLAYBACK METHODS  ////////////////////////////
' Positional sound playback methods will play a sound, depending on the X,Y position of the table element or depending on ActiveBall object position
' These are similar subroutines that are less complicated to use (e.g. simply use standard parameters for the PlaySound call)
' For surround setup - positional sound playback functions will fade between front and rear surround channels and pan between left and right channels
' For stereo setup - positional sound playback functions will only pan between left and right channels
' For mono setup - positional sound playback functions will not pan between left and right channels and will not fade between front and rear channels

' PlaySound full syntax - PlaySound(string, int loopcount, float volume, float pan, float randompitch, int pitch, bool useexisting, bool restart, float front_rear_fade)
' Note - These functions will not work (currently) for walls/slingshots as these do not feature a simple, single X,Y position
Sub PlaySoundAtLevelStatic(playsoundparams, aVol, tableobj)
    PlaySound playsoundparams, 0, aVol * VolumeDial, AudioPan(tableobj), 0, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelExistingStatic(playsoundparams, aVol, tableobj)
    PlaySound playsoundparams, 0, aVol * VolumeDial, AudioPan(tableobj), 0, 0, 1, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelStaticLoop(playsoundparams, aVol, tableobj)
    PlaySound playsoundparams, -1, aVol * VolumeDial, AudioPan(tableobj), 0, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelStaticRandomPitch(playsoundparams, aVol, randomPitch, tableobj)
    PlaySound playsoundparams, 0, aVol * VolumeDial, AudioPan(tableobj), randomPitch, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelActiveBall(playsoundparams, aVol)
	PlaySound playsoundparams, 0, aVol * VolumeDial, AudioPan(ActiveBall), 0, 0, 0, 0, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtLevelExistingActiveBall(playsoundparams, aVol)
	PlaySound playsoundparams, 0, aVol * VolumeDial, AudioPan(ActiveBall), 0, 0, 1, 0, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtLeveTimerActiveBall(playsoundparams, aVol, ballvariable)
	PlaySound playsoundparams, 0, aVol * VolumeDial, AudioPan(ballvariable), 0, 0, 0, 0, AudioFade(ballvariable)
End Sub

Sub PlaySoundAtLevelTimerExistingActiveBall(playsoundparams, aVol, ballvariable)
	PlaySound playsoundparams, 0, aVol * VolumeDial, AudioPan(ballvariable), 0, 0, 1, 0, AudioFade(ballvariable)
End Sub

Sub PlaySoundAtLevelRoll(playsoundparams, aVol, pitch)
    PlaySound playsoundparams, -1, aVol * VolumeDial, AudioPan(tableobj), randomPitch, 0, 0, 0, AudioFade(tableobj)
End Sub


'******************************************************
'				SLINGSHOTS
'******************************************************
Dim BumperSoundFactor
BumperSoundFactor = 15
Dim SlingshotSoundLevel
SlingshotSoundLevel = 0.95												'volume level; range [0, 1]

Dim LStep, RStep

Sub LeftSlingshot_Slingshot
    RandomSoundSlingshotLeft()
	'RandomSoundRubberStrong()
    vpmTimer.PulseSw 49
    LSling.Visible = 0
    LSling1.Visible = 1
    sling2.TransZ = -20
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.TransZ = -10
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.TransZ = 0:LeftSlingShot.TimerEnabled = 0:
    End Select
    LStep = LStep + 1
End Sub


Sub RightSlingshot_Slingshot
    RandomSoundSlingshotRight()
	'RandomSoundRubberStrong()
    vpmTimer.PulseSw 50
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.TransZ = -20
    RStep = 0
    RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.TransZ = -10
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.TransZ = 0:RightSlingShot.TimerEnabled = 0:
    End Select
    RStep = RStep + 1
End Sub


'/////////////////////////////  SLINGSHOT SOLENOID SOUNDS  ////////////////////////////
Sub RandomSoundSlingshotLeft()
	Select Case Int(Rnd*10)+1
		Case 1 : PlaySoundAtLevelStatic SoundFX("TOM_Calle_Sling_Rework_L1_Strong_Layered",DOFContactors), SlingshotSoundLevel, SLING2
		Case 2 : PlaySoundAtLevelStatic SoundFX("TOM_Calle_Sling_Rework_L2_Strong_Layered",DOFContactors), SlingshotSoundLevel, SLING2
		Case 3 : PlaySoundAtLevelStatic SoundFX("TOM_Calle_Sling_Rework_L3_Strong_Layered",DOFContactors), SlingshotSoundLevel, SLING2
		Case 4 : PlaySoundAtLevelStatic SoundFX("TOM_Calle_Sling_Rework_L4_Strong_Layered",DOFContactors), SlingshotSoundLevel, SLING2
		Case 5 : PlaySoundAtLevelStatic SoundFX("TOM_Calle_Sling_Rework_L5_Strong_Layered",DOFContactors), SlingshotSoundLevel, SLING2
		Case 6 : PlaySoundAtLevelStatic SoundFX("TOM_Calle_Sling_Rework_L6_Strong_Layered",DOFContactors), SlingshotSoundLevel, SLING2
		Case 7 : PlaySoundAtLevelStatic SoundFX("TOM_Calle_Sling_Rework_L7_Strong_Layered",DOFContactors), SlingshotSoundLevel, SLING2
		Case 8 : PlaySoundAtLevelStatic SoundFX("TOM_Calle_Sling_Rework_L8_Strong_Layered",DOFContactors), SlingshotSoundLevel, SLING2
		Case 9 : PlaySoundAtLevelStatic SoundFX("TOM_Calle_Sling_Rework_L9_Strong_Layered",DOFContactors), SlingshotSoundLevel, SLING2
		Case 10 : PlaySoundAtLevelStatic SoundFX("TOM_Calle_Sling_Rework_L10_Strong_Layered",DOFContactors), SlingshotSoundLevel, SLING2
	End Select
End Sub

Sub RandomSoundSlingshotRight()
	Select Case Int(Rnd*8)+1
		Case 1 : PlaySoundAtLevelStatic SoundFX("TOM_Calle_Sling_Rework_R1_Strong_Layered",DOFContactors), SlingshotSoundLevel, SLING1
		Case 2 : PlaySoundAtLevelStatic SoundFX("TOM_Calle_Sling_Rework_R2_Strong_Layered",DOFContactors), SlingshotSoundLevel, SLING1
		Case 3 : PlaySoundAtLevelStatic SoundFX("TOM_Calle_Sling_Rework_R3_Strong_Layered",DOFContactors), SlingshotSoundLevel, SLING1
		Case 4 : PlaySoundAtLevelStatic SoundFX("TOM_Calle_Sling_Rework_R4_Strong_Layered",DOFContactors), SlingshotSoundLevel, SLING1
		Case 5 : PlaySoundAtLevelStatic SoundFX("TOM_Calle_Sling_Rework_R5_Strong_Layered",DOFContactors), SlingshotSoundLevel, SLING1
		Case 6 : PlaySoundAtLevelStatic SoundFX("TOM_Calle_Sling_Rework_R6_Strong_Layered",DOFContactors), SlingshotSoundLevel, SLING1
		Case 7 : PlaySoundAtLevelStatic SoundFX("TOM_Calle_Sling_Rework_R7_Strong_Layered",DOFContactors), SlingshotSoundLevel, SLING1
		Case 8 : PlaySoundAtLevelStatic SoundFX("TOM_Calle_Sling_Rework_R9_Strong_Layered",DOFContactors), SlingshotSoundLevel, SLING1
	End Select
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
  'check to make sure that the kickback is disabled and didn't get re-enabled after a kick - from PacDude!

If DesktopMode = True Then
    FadeReel 1 ,L1 'GameOver
    FadeReel 2, L2 'Match Game
end if

    NFadeLm 3, l3
 NFadeLm 3, L3b
If DesktopMode = True Then
    FadeReel 3, L3a 'Shoot Again
end if

    NFadeLm 4, l4
    NFadeLm 4, L4a
	Flash 4, Flasher8
    NFadeLm 5, l5
    NFadeLm 5, L5a

If DesktopMode = True Then
    FadeReel 6, L6 'Ball In Play
end if

    NFadeLm  7, L7
NFadeLm 7, L7a
    NFadeLm 8, L8
NFadeLm 8, L8a
NFadeLm 8, L8b
    NFadeLm 9, L9
	NFadeLm 9, L9a
 NFadeLm 9, L9b
Flash 9, f4
Flash 9, f2
	 NFadeLm 9, L9b1
    NFadeLm 10, L10
NFadeLm 10, L10a
NFadeLm 10, L10b
    NFadeLm 11, L11
NFadeLm 11, L11a
NFadeLm 11, L11b
    NFadeLm 12, l12
NFadeLm 12, L12a
    NFadeLm 13, l13
NFadeLm 13, L13a
	Flash 13, Flasher6
    NFadeLm 14, l14
NFadeLm 14,  L14a
   NFadeLm 105, F105b
	Flash 14, Flasher5
    NFadeLm 15, l15
    NFadeLm 15, L15a


	Flash 15, Flasher4
    NFadeLm 16, l16
    NFadeLm 16, L16a
	Flash 16, Flasher9
    NFadeLm 17, L17
 NFadeLm 17, L17a
	Flash 17, Flasher15
    NFadeLm 18, L18
 NFadeLm 18, L18a
	Flash 18, Flasher16
    NFadeLm 19, L19
 NFadeLm 19, L19a
	Flash 19, Flasher17
    NFadeLm  20, l20
NFadeLm 20, L20a
	NFadeLm 21, L21
 NFadeLm 21, L21a
NFadeLm 21, L21b
    NFadeLm 22, l22
NFadeLm 22, L22a
	Flash 22, Flasher3
    NFadeLm 23, l23
NFadeLm 23, L23a
	Flash 23, Flasher2
    NFadeLm 24, l24
    NFadeLm 24, L24a
	Flash 24, Flasher1
    NFadeLm 25, l25
NFadeLm 25, L25a
	Flash 25, Flasher11
    NFadeLm 26, l26
NFadeLm 26, L26a
	Flash 26, Flasher10
    NFadeLm 27, l27
NFadeLm 27, L27a
	Flash 27, Flasher7
     NFadeLm 28, L28
NFadeLm 28, L28a
	Flash 28, Flasher12
     NFadeLm 29, l29
NFadeLm 29, L29a
	Flash 29, Flasher13
     NFadeLm 30, L30
NFadeLm 30, L30a
	Flash 30, Flasher14
     NFadeLm 31, l31
NFadeLm 31, L31a
     NFadeLm 32, l32
NFadeLm 32, L32a
     NFadeLm 33, l33
NFadeLm 33, L33a
    NFadeLm 34, L34
NFadeLm 34, L34a
     NFadeLm 35, L35
NFadeLm 35, L35a
    NFadeLm 36, L36
 NFadeLm 36, L36a
NFadeLm 36, L36b
    NFadeLm 37, l37
 NFadeLm 37, L37a
    NFadeLm 38, l38
 NFadeLm 38, L38a
   NFadeLm 39, l39
 NFadeLm 39, L39a
    NFadeLm 40, l40
 NFadeLm 40, L40a
     NFadeLm 41, l41
NFadeLm 41, L41a
'   NFadeObjm 42, stoplight_prim, "hslightred copy", "hslightOFF"
'       Flash 42, F42 'Ramp Traffic Light
'   NFadeObjm 43, stoplight_prim, "hslightyellow copy", "hslightOFF"
'       Flash 43, F43 'Ramp Traffic Light
'   NFadeObjm 44, stoplight_prim, "hslightgreen copy", "hslightOFF"
'       Flash 44, F44 'Ramp Traffic Light
    NFadeL 42, l42
    NFadeL 43, l43
    NFadeL 44, l44
    NFadeLm 45, L45
    NFadeLm 45, L45a
    NFadeLm 46, L46
    NFadeLm 46, L46a
    NFadeLm 47, L47
    NFadeLm 47, L47a
    NFadeLm 48, L48
    NFadeLm 48, L48a
    NFadeLm 49, L49
    NFadeLm 49, L49a
    NFadeLm 50, L50
    NFadeLm 50, L50a
    NFadeLm 51, L51
    NFadeLm 51, L51a
    NFadeLm 52, L52
    NFadeLm 52, L52a
    NFadeLm 53, L53
    NFadeLm 53, L53a
    NFadeLm 54, L54
NFadeLm 54, L54a
    NFadeLm 55, L55
NFadeLm 55, L55a
    NFadeLm 56, L56
NFadeLm 56, L56a
    NFadeLm 57, L57
NFadeLm 57, L57a
    NFadeLm 58, L58
NFadeLm 58, L58a
    NFadeLm 59, L59
NFadeLm 59, L59a
    NFadeLm 60, L60
NFadeLm 60, L60a
    NFadeLm 61, L61
NFadeLm 61, L61a
    NFadeLm 62, L62
NFadeLm 62, L62a
    NFadeLm 63, L63
NFadeLm 63, L63a
     NFadeLm 64, L64
NFadeLm 64, L64a

 'Solenoid Controlled Flashers
     NFadeLm 105, F105
   NFadeLm 105, F105b
   NFadeLm 105, F105c
   NFadeLm 105, F105d


   NFadeLm 106, F106
   NFadeLm 106, F106b
   NFadeLm 106, F106c
   NFadeLm 106, F106d

   NFadeLm 109, F109b
   NFadeLm 109, F109c
   NFadeLm 109, f109c1
   NFadeLm 109, F109
   NFadeLm 109, F109a
   NFadeLm 109, f109a1
   NFadeLm 109, f109d
   NFadeLm 109, f109e
NFadeLm 109, f109f
   Flashm 109, Flasherflash2a
Flashm 109, Flasherflash2a1
Flashm 109, Flasherflash2a2
   Flashm 109, Flasherflash1a
   FlupperFlashm 109, Flasherflash2, FlasherLit2, FlasherBase2, F109b
   FlupperFlash 109, Flasherflash1, FlasherLit1, FlasherBase1, F109c


   NFadeLm 112, F112b
   NFadeLm 112, F112c
   NFadeLm 112, f112c1
   NFadeLm 112, F112
   NFadeLm 112, F112a
   NFadeLm 112, f112a1
   NFadeLm 112, f112d
   NFadeLm 112, f112e
   NFadeLm 112, f112f
   Flashm 112, Flasherflash3a
 Flashm 112, Flasherflash4a1
 Flashm 112, Flasherflash4a2
   Flashm 112, Flasherflash4a
   FlupperFlashm 112, Flasherflash4, FlasherLit4, FlasherBase4, F112b
   FlupperFlash 112, Flasherflash3, FlasherLit3, FlasherBase3, F112c

   FlupperFlashm 122, Flasherflash5, FlasherLit5, FlasherBase5, FlasherLight5
   FlupperFlash 122, Flasherflash6, FlasherLit6, FlasherBase6, FlasherLight6
 Flashm 122, Flasherflash8
  Flashm 122, Flasherflash7
Flashm 122, Flasherflash9
  NFadeLm 122, F122e
'Flashm 122, Flasherflash7

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

'**********************************************************************************************************
'Digital Display
'**********************************************************************************************************
 Dim Digits(32)
 Digits(0)=Array(a00, a05, a0c, a0d, a08, a01, a06, a0f, a02, a03, a04, a07, a0b, a0a, a09, a0e)
 Digits(1)=Array(a10, a15, a1c, a1d, a18, a11, a16, a1f, a12, a13, a14, a17, a1b, a1a, a19, a1e)
 Digits(2)=Array(a20, a25, a2c, a2d, a28, a21, a26, a2f, a22, a23, a24, a27, a2b, a2a, a29, a2e)
 Digits(3)=Array(a30, a35, a3c, a3d, a38, a31, a36, a3f, a32, a33, a34, a37, a3b, a3a, a39, a3e)
 Digits(4)=Array(a40, a45, a4c, a4d, a48, a41, a46, a4f, a42, a43, a44, a47, a4b, a4a, a49, a4e)
 Digits(5)=Array(a50, a55, a5c, a5d, a58, a51, a56, a5f, a52, a53, a54, a57, a5b, a5a, a59, a5e)
 Digits(6)=Array(a60, a65, a6c, a6d, a68, a61, a66, a6f, a62, a63, a64, a67, a6b, a6a, a69, a6e)

 Digits(7)=Array(a70, a75, a7c, a7d, a78, a71, a76, a7f, a72, a73, a74, a77, a7b, a7a, a79, a7e)
 Digits(8)=Array(a80, a85, a8c, a8d, a88, a81, a86, a8f, a82, a83, a84, a87, a8b, a8a, a89, a8e)
 Digits(9)=Array(a90, a95, a9c, a9d, a98, a91, a96, a9f, a92, a93, a94, a97, a9b, a9a, a99, a9e)
 Digits(10)=Array(aa0, aa5, aac, aad, aa8, aa1, aa6, aaf, aa2, aa3, aa4, aa7, aab, aaa, aa9, aae)
 Digits(11)=Array(ab0, ab5, abc, abd, ab8, ab1, ab6, abf, ab2, ab3, ab4, ab7, abb, aba, ab9, abe)
 Digits(12)=Array(ac0, ac5, acc, acd, ac8, ac1, ac6, acf, ac2, ac3, ac4, ac7, acb, aca, ac9, ace)
 Digits(13)=Array(ad0, ad5, adc, add, ad8, ad1, ad6, adf, ad2, ad3, ad4, ad7, adb, ada, ad9, ade)

 ' 3rd Player
Digits(14) = Array(LED150,LED151,LED152,LED153,LED154,LED155,LED156)
Digits(15) = Array(LED160,LED161,LED162,LED163,LED164,LED165,LED166)
Digits(16) = Array(LED170,LED171,LED172,LED173,LED174,LED175,LED176)
Digits(17) = Array(LED180,LED181,LED182,LED183,LED184,LED185,LED186)
Digits(18) = Array(LED190,LED191,LED192,LED193,LED194,LED195,LED196)
Digits(19) = Array(LED200,LED201,LED202,LED203,LED204,LED205,LED206)
Digits(20) = Array(LED210,LED211,LED212,LED213,LED214,LED215,LED216)

' 4th Player
Digits(21) = Array(LED220,LED221,LED222,LED223,LED224,LED225,LED226)
Digits(22) = Array(LED230,LED231,LED232,LED233,LED234,LED235,LED236)
Digits(23) = Array(LED240,LED241,LED242,LED243,LED244,LED245,LED246)
Digits(24) = Array(LED250,LED251,LED252,LED253,LED254,LED255,LED256)
Digits(25) = Array(LED260,LED261,LED262,LED263,LED264,LED265,LED266)
Digits(26) = Array(LED270,LED271,LED272,LED273,LED274,LED275,LED276)
Digits(27) = Array(LED280,LED281,LED282,LED283,LED284,LED285,LED286)

' Ball in play
Digits(28) = Array(LED290,LED291,LED292,LED293,LED294,LED295,LED296)
Digits(29) = Array(LED300,LED301,LED302,LED303,LED304,LED305,LED306)

' Num of Credits
Digits(30) = Array(LED310,LED311,LED312,LED313,LED314,LED315,LED316)
Digits(31) = Array(LED320,LED321,LED322,LED323,LED324,LED325,LED326)

 Sub DisplayTimer_Timer
    Dim ChgLED, ii, jj, num, chg, stat, obj, b, x
    ChgLED=Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
    If Not IsEmpty(ChgLED)Then
        If DesktopMode = True Then
       For ii=0 To UBound(chgLED)
          num=chgLED(ii, 0) : chg=chgLED(ii, 1) : stat=chgLED(ii, 2)
            if (num < 32) then
              For Each obj In Digits(num)
                   If chg And 1 Then obj.State=stat And 1
                   chg=chg\2 : stat=stat\2
                  Next
            Else
                   end if
        Next
       end if
    End If
 End Sub


' *** JP Fading Lights adaptation of Flupper's flasher scripting **


Sub FlupperFlash(nr, FlashObject, LitObject, BaseObject, LightObject)
    FlupperFlashm nr, FlashObject, LitObject, BaseObject, LightObject
    FadeEmpty nr
End Sub

Sub FlupperFlashm(nr, FlashObject, LitObject, BaseObject, LightObject)
    'exit sub
    dim flashx3
    Select Case FadingLevel(nr)
        Case 4, 5
            ' This section adapted from Flupper's script
            flashx3 = FlashLevel(nr) * FlashLevel(nr) * FlashLevel(nr)
            FlashObject.IntensityScale = flashx3
            LitObject.BlendDisableLighting = 10 * flashx3
            BaseObject.BlendDisableLighting = flashx3
            LightObject.IntensityScale = flashx3
            LitObject.material = "domelit" & Round(9 * FlashLevel(nr))
            LitObject.visible = 1
            FlashObject.visible = 1
        case 3:
            LitObject.visible = 0
            FlashObject.visible = 0
    end select
End Sub

Sub FadeEmpty(nr)   'Fade a lamp number, no object updates
    Select Case FadingLevel(nr)
        Case 3
            FadingLevel(nr) = 0
        Case 4 'off
            FlashLevel(nr) = FlashLevel(nr) - FlashSpeedDown(nr)
            If FlashLevel(nr) < FlashMin(nr) Then
                FlashLevel(nr) = FlashMin(nr)
               FadingLevel(nr) = 3 'completely off
            End if
            'Object.IntensityScale = FlashLevel(nr)
        Case 5 ' on
            FlashLevel(nr) = FlashLevel(nr) + FlashSpeedUp(nr)
            If FlashLevel(nr) > FlashMax(nr) Then
                FlashLevel(nr) = FlashMax(nr)
                FadingLevel(nr) = 6 'completely on
            End if
            'Object.IntensityScale = FlashLevel(nr)
        Case 6
            FadingLevel(nr) = 1
    End Select
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
        If BallVel(BOT(b) ) > 1 AND BOT(b).z < 30 Then
            rolling(b) = True
            PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) ), Pan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0, AudioFade(BOT(b))
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
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 800, Pan(ball1), 0, Pitch(ball1), 0, 0
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
    PlaySound "pinhit_low", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Targets_Hit (idx)
    PlaySound "target", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Metals_Thin_Hit (idx)
    PlaySound "metalhit_medium", 0, 3000*Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Metals_Medium_Hit (idx)
    PlaySound "metalhit_medium", 0, 3000*Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Metals2_Hit (idx)
    PlaySound "metalhit_thin", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Metals3_Hit (idx)
    PlaySound "metalhit_thin", 0, 11000*Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Gates_Hit (idx)
    PlaySound "gate4", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Spinner_Spin
    PlaySound "fx_spinner",0,.25,0,0.25
End Sub

'/////////////////////////////  VARIOUS PLAYFIELD SOUND SUBROUTINES  ////////////////////////////
'/////////////////////////////  RUBBERS AND POSTS  ////////////////////////////
'/////////////////////////////  RUBBERS - EVENTS  ////////////////////////////
Sub Rubbers_Hit(idx)
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 5 then
 		RandomSoundRubberStrong()
	End if
	If finalspeed <= 5 then
 		RandomSoundRubberWeak()
 	End If
End Sub


'Sub Rubbers_Hit(idx)
'    dim finalspeed
'    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
'    If finalspeed > 20 then
'        PlaySound "fx_rubber_band", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
'    End if
'    If finalspeed >= 6 AND finalspeed <= 20 then
'        RandomSoundRubber()

'    End If
'End Sub

'/////////////////////////////  POSTS - EVENTS  ////////////////////////////
Sub Posts_Hit(idx)
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 5 then
 		RandomSoundRubberStrong()
	End if
	If finalspeed <= 5 then
 		RandomSoundRubberWeak()
 	End If
End Sub


'Sub Posts_Hit(idx)
'    dim finalspeed
'    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
'    If finalspeed > 16 then
'        PlaySound "fx_postrubber", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
'    End if
'    If finalspeed >= 6 AND finalspeed <= 16 then
'        RandomSoundRubber()
'    End If
'End Sub

'Sub RandomSoundRubber()
'    Select Case Int(Rnd*3)+1
'        Case 1 : PlaySound "fx_rubber_band", 0, 80*Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
'        Case 2 : PlaySound "fx_rubber_band", 0, 80*Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
'        Case 3 : PlaySound "fx_rubber_band", 0, 80*Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
'    End Select
'End Sub

FlipperUpAttackMinimumSoundLevel = 0.010           						'volume level; range [0, 1]
FlipperUpAttackMaximumSoundLevel = 0.635								'volume level; range [0, 1]
FlipperUpSoundLevel = 0.5                       						'volume level; range [0, 1]
FlipperDownSoundLevel = 0.25                      						'volume level; range [0, 1]
FlipperLeftHitParm = FlipperUpSoundLevel
FlipperRightHitParm = FlipperUpSoundLevel								'sound helper; not configurable


'/////////////////////////////  FLIPPER BATS SOLENOID CORE SOUND  ////////////////////////////
Sub RandomSoundFlipperUpLeft()
	Select Case Int(Rnd*11)+1
		Case 1 : PlaySoundAtLevelStatic SoundFX("TOM_Calle_Flipper_L01",DOFFlippers), FlipperLeftHitParm, LeftFlipper
		Case 2 : PlaySoundAtLevelStatic SoundFX("TOM_Calle_Flipper_L02",DOFFlippers), FlipperLeftHitParm, LeftFlipper
		Case 3 : PlaySoundAtLevelStatic SoundFX("TOM_Calle_Flipper_L07",DOFFlippers), FlipperLeftHitParm, LeftFlipper
		Case 4 : PlaySoundAtLevelStatic SoundFX("TOM_Calle_Flipper_L08",DOFFlippers), FlipperLeftHitParm, LeftFlipper
		Case 5 : PlaySoundAtLevelStatic SoundFX("TOM_Calle_Flipper_L09",DOFFlippers), FlipperLeftHitParm, LeftFlipper
		Case 6 : PlaySoundAtLevelStatic SoundFX("TOM_Calle_Flipper_L10",DOFFlippers), FlipperLeftHitParm, LeftFlipper
		Case 7 : PlaySoundAtLevelStatic SoundFX("TOM_Calle_Flipper_L12",DOFFlippers), FlipperLeftHitParm, LeftFlipper
		Case 8 : PlaySoundAtLevelStatic SoundFX("TOM_Calle_Flipper_L14",DOFFlippers), FlipperLeftHitParm, LeftFlipper
		Case 9 : PlaySoundAtLevelStatic SoundFX("TOM_Calle_Flipper_L18",DOFFlippers), FlipperLeftHitParm, LeftFlipper
		Case 10 : PlaySoundAtLevelStatic SoundFX("TOM_Calle_Flipper_L20",DOFFlippers), FlipperLeftHitParm, LeftFlipper
		Case 11 : PlaySoundAtLevelStatic SoundFX("TOM_Calle_Flipper_L26",DOFFlippers), FlipperLeftHitParm, LeftFlipper
	End Select
End Sub

Sub RandomSoundFlipperUpRight()
	Select Case Int(Rnd*11)+1
		Case 1 : PlaySoundAtLevelStatic SoundFX("TOM_Calle_Flipper_R01",DOFFlippers), FlipperRightHitParm, RightFlipper
		Case 2 : PlaySoundAtLevelStatic SoundFX("TOM_Calle_Flipper_R02",DOFFlippers), FlipperRightHitParm, RightFlipper
		Case 3 : PlaySoundAtLevelStatic SoundFX("TOM_Calle_Flipper_R03",DOFFlippers), FlipperRightHitParm, RightFlipper
		Case 4 : PlaySoundAtLevelStatic SoundFX("TOM_Calle_Flipper_R04",DOFFlippers), FlipperRightHitParm, RightFlipper
		Case 5 : PlaySoundAtLevelStatic SoundFX("TOM_Calle_Flipper_R05",DOFFlippers), FlipperRightHitParm, RightFlipper
		Case 6 : PlaySoundAtLevelStatic SoundFX("TOM_Calle_Flipper_R06",DOFFlippers), FlipperRightHitParm, RightFlipper
		Case 7 : PlaySoundAtLevelStatic SoundFX("TOM_Calle_Flipper_R07",DOFFlippers), FlipperRightHitParm, RightFlipper
		Case 8 : PlaySoundAtLevelStatic SoundFX("TOM_Calle_Flipper_R08",DOFFlippers), FlipperRightHitParm, RightFlipper
		Case 9 : PlaySoundAtLevelStatic SoundFX("TOM_Calle_Flipper_R09",DOFFlippers), FlipperRightHitParm, RightFlipper
		Case 10 : PlaySoundAtLevelStatic SoundFX("TOM_Calle_Flipper_R10",DOFFlippers), FlipperRightHitParm, RightFlipper
		Case 11 : PlaySoundAtLevelStatic SoundFX("TOM_Calle_Flipper_R11",DOFFlippers), FlipperRightHitParm, RightFlipper
	End Select
End Sub


Sub RandomSoundFlipperDownLeft()
	Select Case Int(Rnd*7)+1
		Case 1 : PlaySoundAtLevelStatic SoundFX("WD_TOM_Flipper_Left_Down_1_Bass",DOFFlippers), FlipperDownSoundLevel, LeftFlipper
		Case 2 : PlaySoundAtLevelStatic SoundFX("WD_TOM_Flipper_Left_Down_2_Bass",DOFFlippers), FlipperDownSoundLevel, LeftFlipper
		Case 3 : PlaySoundAtLevelStatic SoundFX("WD_TOM_Flipper_Left_Down_3_Bass",DOFFlippers), FlipperDownSoundLevel, LeftFlipper
		Case 4 : PlaySoundAtLevelStatic SoundFX("WD_TOM_Flipper_Left_Down_4_Bass",DOFFlippers), FlipperDownSoundLevel, LeftFlipper
		Case 5 : PlaySoundAtLevelStatic SoundFX("WD_TOM_Flipper_Left_Down_5_Bass",DOFFlippers), FlipperDownSoundLevel, LeftFlipper
		Case 6 : PlaySoundAtLevelStatic SoundFX("WD_TOM_Flipper_Left_Down_6_Bass",DOFFlippers), FlipperDownSoundLevel, LeftFlipper
		Case 7 : PlaySoundAtLevelStatic SoundFX("WD_TOM_Flipper_Left_Down_7_Bass",DOFFlippers), FlipperDownSoundLevel, LeftFlipper
	End Select
End Sub

Sub RandomSoundFlipperDownRight()
	Select Case Int(Rnd*8)+1
		Case 1 : PlaySoundAtLevelStatic SoundFX("WD_TOM_Flipper_Right_Down_1_Bass",DOFFlippers), FlipperDownSoundLevel, RightFlipper
		Case 2 : PlaySoundAtLevelStatic SoundFX("WD_TOM_Flipper_Right_Down_2_Bass",DOFFlippers), FlipperDownSoundLevel, RightFlipper
		Case 3 : PlaySoundAtLevelStatic SoundFX("WD_TOM_Flipper_Right_Down_3_Bass",DOFFlippers), FlipperDownSoundLevel, RightFlipper
		Case 4 : PlaySoundAtLevelStatic SoundFX("WD_TOM_Flipper_Right_Down_4_Bass",DOFFlippers), FlipperDownSoundLevel, RightFlipper
		Case 5 : PlaySoundAtLevelStatic SoundFX("WD_TOM_Flipper_Right_Down_5_Bass",DOFFlippers), FlipperDownSoundLevel, RightFlipper
		Case 6 : PlaySoundAtLevelStatic SoundFX("WD_TOM_Flipper_Right_Down_6_Bass",DOFFlippers), FlipperDownSoundLevel, RightFlipper
		Case 7 : PlaySoundAtLevelStatic SoundFX("WD_TOM_Flipper_Right_Down_7_Bass",DOFFlippers), FlipperDownSoundLevel, RightFlipper
		Case 8 : PlaySoundAtLevelStatic SoundFX("WD_TOM_Flipper_Right_Down_8_Bass",DOFFlippers), FlipperDownSoundLevel, RightFlipper
	End Select
End Sub


'/////////////////////////////  FLIPPER BATS BALL COLLIDE SOUND  ////////////////////////////
dim angdamp, veldamp
angdamp = 0.2
veldamp = 0.3

Sub LeftFlipper_Collide(parm)
	If parm > 4  and leftflipper.currentangle < 90 Then
		activeball.angmomx=activeball.angmomx*angdamp
		activeball.angmomy=activeball.angmomy*angdamp
		activeball.angmomz=activeball.angmomz*angdamp
		If  activeball.velx > 0 Then activeball.velx = activeball.velx * veldamp
	End If

	FlipperLeftHitParm = parm/10
	If FlipperLeftHitParm > 1 Then
		FlipperLeftHitParm = 1
	End If
	FlipperLeftHitParm = FlipperUpSoundLevel * FlipperLeftHitParm
	RandomSoundRubberFlipper(parm)
End Sub

Sub RightFlipper_Collide(parm)
	If parm > 4  and rightflipper.currentangle > -90 Then
		activeball.angmomx=activeball.angmomx*angdamp
		activeball.angmomy=activeball.angmomy*angdamp
		activeball.angmomz=activeball.angmomz*angdamp
		If  activeball.velx < 0 Then activeball.velx = activeball.velx * veldamp
	End If

	FlipperRightHitParm = parm/10
	If FlipperRightHitParm > 1 Then
		FlipperRightHitParm = 1
	End If
	FlipperRightHitParm = FlipperUpSoundLevel * FlipperRightHitParm
 	RandomSoundRubberFlipper(parm)
End Sub

Sub LeftFlipper_Collide(parm)
    RandomSoundRubberFlipper(parm)
End Sub

Sub RightFlipper_Collide(parm)
    RandomSoundRubberFlipper(parm)
End Sub


Sub RandomSoundRubberFlipper(parm)
	Select Case Int(Rnd*7)+1
		Case 1 : PlaySoundAtLevelActiveBall ("TOM_Rubber_Flipper_Normal_1"), parm / 10 * RubberFlipperSoundFactor
		Case 2 : PlaySoundAtLevelActiveBall ("TOM_Rubber_Flipper_Normal_2"), parm / 10 * RubberFlipperSoundFactor
		Case 3 : PlaySoundAtLevelActiveBall ("TOM_Rubber_Flipper_Normal_3"), parm / 10 * RubberFlipperSoundFactor
		Case 4 : PlaySoundAtLevelActiveBall ("TOM_Rubber_Flipper_Normal_4"), parm / 10 * RubberFlipperSoundFactor
		Case 5 : PlaySoundAtLevelActiveBall ("TOM_Rubber_Flipper_Normal_5"), parm / 10 * RubberFlipperSoundFactor
		Case 6 : PlaySoundAtLevelActiveBall ("TOM_Rubber_Flipper_Normal_6"), parm / 10 * RubberFlipperSoundFactor
		Case 7 : PlaySoundAtLevelActiveBall ("TOM_Rubber_Flipper_Normal_7"), parm / 10 * RubberFlipperSoundFactor
	End Select
End Sub


'Sub RandomSoundFlipper()
'    Select Case Int(Rnd*3)+1
'        Case 1 : PlaySound "fx_rubber_flipper", 0, 160*Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
'        Case 2 : PlaySound "fx_rubber_flipper", 0, 160*Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
'        Case 3 : PlaySound "fx_rubber_flipper", 0, 160*Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
'    End Select
'End Sub


'*********** BALL SHADOW *********************************
Dim BallShadow
BallShadow = Array (BallShadow1, BallShadow2, BallShadow3)

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

 '*********** FLUPPER'S BATS *********************************
TimerFlipper.interval = 10
TimerFlipper.Enabled = 1
Sub TimerFlipper_Timer()
    batleft.objrotz = LeftFlipper.CurrentAngle + 1
    batleftshadow.objrotz = batleft.objrotz

    batright.objrotz = RightFlipper.CurrentAngle + 1
    batrightshadow.objrotz = batright.objrotz
    batright1.objrotz = RightFlipper1.CurrentAngle + 1
    batrightshadow.objrotz = batright.objrotz
    If l42.state=1 and l43.state=0 and l44.state=0 then stoplight_prim.image="hslightred copy"
    If l42.state=0 and l43.state=1 and l44.state=0 then stoplight_prim.image="hslightyellow copy"
    If l42.state=0 and l43.state=0 and l44.state=1 then stoplight_prim.image="hslightgreen copy"
    If l42.state=0 and l43.state=1 and l44.state=1 then stoplight_prim.image="hslightgreenyellow copy"
    If l42.state=0 and l43.state=0 and l44.state=0 then stoplight_prim.image="hslightoff"
    If l42.state=1 and l43.state=1 and l44.state=1 then stoplight_prim.image="hslighton"
End sub

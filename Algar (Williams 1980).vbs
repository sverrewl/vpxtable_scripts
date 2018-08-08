Option Explicit
Randomize

' Special thanks goes to previous authors of this tables : Practicedummy, UncleWilly, LuvThatApex and Kristian.
' Background image for Desktop by Batch. Flippers by Flupper. A few primitives from GtxJoe's primitive collection.
' Main table by Thalamus and playfield done mostly by Kalavera ( thank you so much )
' The biggest thanks goes to the main devs - without you this would not be possible. You guys rock !
' The table started out from the example table and there is code and resources there provided by the community.
' JP, Ninuzzu, DjRobX, Rothbauerw probably also 32assassins. Thanks guys !

' Thalamus 2018-07-24
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Changed UseSolenoids=1 to 2
' No special SSF tweaks yet.
' Wob 2018-08-08
' Added vpmInit Me to table init and both cSingleLFlip and /cSingleRFlip

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Const cGameName="algar_l1",UseSolenoids=2,UseLamps=1,UseGI=0,SSolenoidOn="SolOn",SSolenoidOff="SolOff",SFlipperOn="FlipperUp",SFlipperOff="FlipperDown"
Const SCoin="coin3",cCredits="Algar"
' Wob: Added for Fast Flips (No upper Flippers)
Const cSingleLFlip = 0
Const cSingleRFlip = 0

LoadVPM "01520000","s4.vbs",3.1

Dim DesktopMode: DesktopMode = Table1.ShowDT

If DesktopMode = True Then 'Show Desktop components
  lrail.visible=1
  rrail.visible=1
  SideCab.visible=1
  Backwall.visible=1
  Screw1.visible=1
  Screw2.visible=1
  Screw3.visible=1
  Screw4.visible=1
  Screw5.visible=1
Else
  lrail.visible=0
  rrail.visible=0
  SideCab.visible=0
  Backwall.visible=0
  Screw1.visible=0
  Screw2.visible=0
  Screw3.visible=0
  Screw4.visible=0
  Screw5.visible=0
End if

Dim bstrough,bslow,bshig
Dim dtbank3,dtbank5,i
Dim Bumper1AnimationCount
Dim Bumper2AnimationCount
Dim Bumper3AnimationCount
Dim EnableBallControl
Dim ShowBallShadow
Dim FlipperShadows
Dim LMPFEnabled

' Options

FlipperShadows = 1  	' 1 turns on, 0 turns off flipper shadows.
ShowBallShadow = 0  	' 1 turns on, 0 turns off ball shadows
LMPFEnabled    = 1		' 1 turns on, 0 turns off Lightmap

' You need to add some extra elements if you want BallControl - I just left in the code required.
EnableBallControl = False 	' Change to true or 1 to enable manual ball control (or press C in-game) via the arrow keys and B (boost movement) keys


Sub Plunger_Init()
	PlaySound SoundFX("ballrelease",DOFContactors), 0,1,AudioPan(BallRelease),0.25,0,0,1,AudioFade(BallRelease)
End Sub

Sub Table1_KeyDown(ByVal keycode)
	If keycode = PlungerKey Then
		Plunger.PullBack
		PlaySound "plungerpull",0,1,AudioPan(Plunger),0.25,0,0,1,AudioFade(Plunger)
	End If

	If keycode = LeftTiltKey Then
		Nudge 90, 2
	End If

	If keycode = RightTiltKey Then
		Nudge 270, 2
	End If

	If keycode = CenterTiltKey Then
		Nudge 0, 2
	End If

    ' Manual Ball Control
	If keycode = 46 Then	 				' C Key
		If EnableBallControl = 1 Then
			EnableBallControl = 0
		Else
			EnableBallControl = 1
		End If
	End If
    If EnableBallControl = 1 Then
		If keycode = 48 Then 				' B Key
			If BCboost = 1 Then
				BCboost = BCboostmulti
			Else
				BCboost = 1
			End If
		End If
		If keycode = 203 Then BCleft = 1	' Left Arrow
		If keycode = 200 Then BCup = 1		' Up Arrow
		If keycode = 208 Then BCdown = 1	' Down Arrow
		If keycode = 205 Then BCright = 1	' Right Arrow
	End If
	if vpmKeyDown(keycode) Then Exit Sub
End Sub

Sub Table1_KeyUp(ByVal keycode)
	If keycode = PlungerKey Then
		Plunger.Fire
		PlaySound "plunger",0,1,AudioPan(Plunger),0.25,0,0,1,AudioFade(Plunger)
	End If

    'Manual Ball Control
	If EnableBallControl = 1 Then
		If keycode = 203 Then BCleft = 0	' Left Arrow
		If keycode = 200 Then BCup = 0		' Up Arrow
		If keycode = 208 Then BCdown = 0	' Down Arrow
		If keycode = 205 Then BCright = 0	' Right Arrow
	End If
	if vpmKeyUp(keycode) Then Exit Sub
End Sub


' SolCallbacks

SolCallback(1)="solballrelease"
SolCallback(2)="bslow.solout"
SolCallback(3)="dtbank3.soldropup"
SolCallback(4)="dtbank5.soldropup"
SolCallback(5)="bshig.solout"
' SolCallback(6)= "vpmSolDiverter OutlaneGate, solon, Not"
SolCallback(6)="SolDiverterSub"
SolCallback(7)="chamber"
SolCallback(14)="vpmsolsound SoundFX(""knocker"",DOFKnocker),"
' SolCallback(sllflipper)="vpmsolflipper leftflipper,nothing,"
' SolCallback(slrflipper)="vpmsolflipper rightflipper,nothing,"
SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"
' SolCallback(32) = "VpmNudge.SolGameOn"
SolCallback(23) = "VpmNudge.SolGameOn"


Sub SolDiverterSub(Enabled)
  If Enabled Then
    OutlaneGate.RotateToStart
	Light37.State = LightStateOn
    PlaySound SoundFX(SSolenoidOn,DOFContactors)
  Else
    OutlaneGate.RotateToEnd
	Light37.State = LightStateOff
  End If
End Sub

Sub SolLFlipper(Enabled)
     If Enabled Then
		LeftFlipper.RotateToEnd
'		PlaySound SoundFX("fx_flipperup",DOFFlippers), 0, 2, AudioPan(LeftFlipper), 0.05,0,0,1,AudioFade(LeftFlipper)
		PlaySound SoundFX("",DOFFlippers), 0, 2, AudioPan(LeftFlipper), 0.05,0,0,1,AudioFade(LeftFlipper)
		PlaySoundAt "fx_flipperup",LeftFlipper
     Else
		 LeftFlipper.RotateToStart
		 PlaySound SoundFX("fx_flipperdown",DOFFlippers), 0, 1, AudioPan(LeftFlipper), 0.05,0,0,1,AudioFade(LeftFlipper)
		 PlaySound SoundFX("",DOFFlippers), 0, 1, AudioPan(LeftFlipper), 0.05,0,0,1,AudioFade(LeftFlipper)
		 PlaySoundAt "fx_flipperdown",LeftFlipper
     End If
  End Sub

Sub SolRFlipper(Enabled)
     If Enabled Then
		RightFlipper.RotateToEnd
'		PlaySound SoundFX("fx_flipperup",DOFFlippers), 0, 2, AudioPan(RightFlipper), 0.05,0,0,1,AudioFade(RightFlipper)
		PlaySound SoundFX("",DOFFlippers), 0, 2, AudioPan(RightFlipper), 0.05,0,0,1,AudioFade(RightFlipper)
		PlaySoundAt "fx_flipperup",RightFlipper
		controller.switch(50)=false
     Else
		RightFlipper.RotateToStart
'		PlaySound SoundFX("fx_flipperdown",DOFFlippers), 0, 1, AudioPan(RightFlipper), 0.05,0,0,1,AudioFade(RightFlipper)
		PlaySound SoundFX("",DOFFlippers), 0, 1, AudioPan(RightFlipper), 0.05,0,0,1,AudioFade(RightFlipper)
		PlaySoundAt "fx_flipperdown",RightFlipper
		controller.switch(50)=true
     End If
End Sub

Sub Table1_Init()
    vpmInit Me
    With Controller
       .GameName = cGameName
        If Err Then MsgBox"Can't start Game"&cGameName&vbNewLine&Err.Description:Exit Sub
       .SplashInfoLine = "Algar, Williams 1980. Table playfield recreation by VPX by Thalamus & Kalavera"
        .HandleMechanics=0
        .HandleKeyboard=0
        .ShowDMDOnly=1
        .ShowFrame=0
        .ShowTitle=0
        .Hidden=0
       On Error Resume Next
        .Run GetPlayerHwnd
        If Err Then MsgBox Err.Description
    End With

   PinMAMETimer.Interval=PinMAMEInterval:PinMAMETimer.Enabled=1

    Set bsTrough=New cvpmBallStack
    bsTrough.InitNoTrough BallRelease,9,80,2
    bsTrough.InitExitSnd SoundFX("BallRelease",DOFContactors),SoundFX("SolOn",DOFContactors)

    Set bslow =new cvpmballstack
    bslow.initsaucer kicker2,19,5,17 ' 16
    bslow.kickanglevar=5
    bslow.initexitsnd SoundFX("diverter",DOFContactors), ","

    set bshig =new cvpmballstack
    bshig.initsaucer kicker1,26,175,4
    bshig.kickanglevar=5
    bslow.initexitsnd SoundFX("rsling",DOFContactors), ","

    set dtbank3=new cvpmdroptarget
    dtbank3.initdrop array(DropTarget1,DropTarget2,DropTarget3),array(14,15,16)
    dtbank3.initsnd SoundFX("dtl",DOFContactors), SoundFX("dts",DOFContactors)
    dtbank3.AllDownSw = 17

    set dtbank5=new cvpmdroptarget
    dtbank5.initdrop array (DropTarget4,DropTarget5,DropTarget6),array(39,40,41)
    dtbank5.initsnd SoundFX("dtl",DOFContactors), SoundFX("dts",DOFContactors)
    dtbank5.AllDownSw = 42

' I wonder why tilt isn't working for bumpers and slings ? Or, now it is when I stole the SolCallBack # from Gorgar - but, I can't find it in the doc ?!
' Seems using both tiltswitch=51 and 1 is working. Again why ?

 	vpmnudge.tiltswitch=51
'	vpmNudge.TiltSwitch=1
    vpmnudge.sensitivity=5
    vpmNudge.TiltObj = Array(Bumper1,Bumper2,Bumper3,LeftSlingshot,RightSlingshot,URightSlingshot)

	If ShowBallShadow = 1 Then
		BallShadowUpdate.enabled=1
		GI9.state=1
	Else
		BallShadowUpdate.enabled=0
		GI9.state=0
	End If

	If FlipperShadows = 1 then
        FlipperLSh.visible=1
        FlipperRSh.visible=1
       else
        FlipperLSh.visible=0
        FlipperRSh.visible=0
    End If

	Sw11.IsDropped = True
	Sw12.IsDropped = True
	Sw13.IsDropped = True
	Sw24.IsDropped = True
	Sw25.IsDropped = True
	Sw27.IsDropped = True
	Sw28.IsDropped = True
	Sw35.IsDropped = True
	Sw37.IsDropped = True
	Sw44.IsDropped = True


' Create Captive Ball

	RCaptKicker1.CreateBall
	RCaptKicker1.Kick 180,10
	RCaptKicker2.CreateBall
	RCaptKicker2.Kick 180,10
	RCaptKicker3.CreateBall
	RCaptKicker3.Kick 180,10

 End Sub

' Bumpers

Sub Bumper1_Hit
	PlaySound SoundFX("fx_bumper1",DOFContactors), 0,1,AudioPan(Bumper1),0,0,0,1,AudioFade(Bumper1)
'	PlaySound SoundFX("",DOFContactors), 0,1,AudioPan(Bumper1),0,0,0,1,AudioFade(Bumper1)
'	PlaySoundAt "fx_bumper1",Bumper1
	vpmtimer.pulsesw 47
	Me.TimerEnabled = 1
End Sub

Sub Bumper1_Timer
  Me.Timerenabled = 0
End Sub

Sub Bumper2_Hit
  	PlaySound SoundFX("fx_bumper2",DOFContactors), 0,1,AudioPan(Bumper2),0,0,0,1,AudioFade(Bumper2)
'	PlaySound SoundFX("",DOFContactors), 0,1,AudioPan(Bumper2),0,0,0,1,AudioFade(Bumper2)
'	PlaySoundAt "fx_bumper2",Bumper2
	vpmtimer.pulsesw 48
  Me.TimerEnabled = 1
End Sub

Sub Bumper2_Timer
  Me.Timerenabled = 0
End Sub

Sub Bumper3_Hit
  	PlaySound SoundFX("fx_bumper3",DOFContactors), 0,1,AudioPan(Bumper3),0,0,0,1,AudioFade(Bumper3)
'	PlaySound SoundFX("",DOFContactors), 0,1,AudioPan(Bumper3),0,0,0,1,AudioFade(Bumper3)
'	PlaySoundAt "fx_bumper3",Bumper3
	vpmtimer.pulsesw 49
  Me.TimerEnabled = 1
End Sub

Sub Bumper3_Timer
	Me.Timerenabled = 0
End Sub


 Sub solballrelease(enabled)
     bstrough.solexit SoundFX(ssolenoidon,DOFContactors), SoundFX(ssolenoidon,DOFContactors),enabled
 End Sub

Sub DrainSound_Hit()
	PlaySound "drain",0,1,AudioPan(Drain),0.25,0,0,1,AudioFade(Drain)
End Sub

 Sub Drain_Hit()
   bstrough.addball me
 End Sub

sub chamber(enabled)
  'light1000.state=1:timer1.enabled=1
  PlaySound "SolOff"
  CapturedBallPost1.IsDropped=1
  CapturedBallPost2.IsDropped=1
  CapturedBallPost3.IsDropped=1
end sub

'Sub Timer1_Timer()
'  light1000.state=0:timer1.enabled=0
'End Sub

Set LampCallback = GetRef("UpdateLamps")

Sub UpdateLamps
	Light5a.State=Light5.State
	Light4f.State=Light4.State
	Light9f.State=Light9.State
	Light10f.State=Light10.State
	Light11f.State=Light11.State
	Light12f.State=Light12.State
	Light13f.State=Light13.State
	Light14f.State=Light14.State
	Light21f.State=Light21.State
	Light22f.State=Light22.State
	Light23f.State=Light23.State
	Light24f.State=Light24.State
End Sub

 set lights(1)=light1 ' same player shoot again
 set lights(2)=light2 ' left special
 set lights(3)=light3 ' right special
 set lights(4)=light4 ' loop gate x2
 set lights(5)=light5 ' chamber 50k
 set lights(6)=light6 ' chamber 40k
 set lights(7)=light7 ' chamber 30k
 set lights(8)=light8 ' chamber reset
 set lights(9)=light9 ' loop lane 10
 set lights(10)=light10 ' loop lane 20
 set lights(11)=light11 ' loop lane 30
 set lights(12)=light12 ' loop lane 40
 set lights(13)=light13 ' loop lane 50
 set lights(14)=light14 ' loop lane 60
 set lights(15)=light15 ' extra ball when lit
 'set lights(16)=light16 ' was off - says not used but why the hell not ?
 set lights(17)=light17 ' 2x
 set lights(18)=light18 ' 3x
 set lights(19)=light19 ' 4x
 set lights(20)=light20 ' 5x
 set lights(21)=light21 ' K rollover
 set lights(22)=light22 ' O rollover
 set lights(23)=light23 ' R rollover
 set lights(24)=light24 ' A rollover
 set lights(25)=light25 ' left 3 bank left target arrow
 set lights(26)=light26 ' left 3 bank center target arrow
 set lights(27)=light27 ' left 3 bank right target arrow
 set lights(28)=light28 ' center 3 bank left target arrow
 set lights(29)=light29 ' center 3 bank center target arrow
 set lights(30)=light30 ' center 3 bank right target arrow
 set lights(31)=light31 ' left spinner
 set lights(32)=light32 ' right spinner
 set lights(33)=light33 ' 3 banks 10 bonus
 set lights(34)=light34 ' 3 banks 30 bonus
 set lights(35)=light35 ' 3 banks 50 bonus
 set lights(36)=light36 ' 3 banks 100 bonus
 ' set lights(37)=light37 ' was off - doc say, not used
 set lights(38)=light38 ' 20k bonus
 set lights(39)=light39 ' 10k bonus
 set lights(40)=light40 ' 1k bonus
 set lights(41)=light41 ' 2k bonus
 set lights(42)=light42 ' 3k bonus
 set lights(43)=light43 ' 4k bonus
 set lights(44)=light44 ' 5k bonus
 set lights(45)=light45 ' 6k bonus
 set lights(46)=light46 ' 7k bonus
 set lights(47)=light47 ' 8k bonus
 set lights(48)=light48 ' 9k bonus
 set lights(49)=light49 ' was off - doc say, not used
 set lights(50)=light50 ' was off - 1 can play
 set lights(51)=light51 ' was off - 2 can play
 set lights(52)=light52 ' was off - 3 can play
 set lights(53)=light53 ' was off - 4 can play
 set lights(54)=light54 ' was off - match
 set lights(55)=light55 ' ball in play
 set lights(56)=light56 ' credits on playfield
 set lights(57)=light57 ' 1 player up
 set lights(58)=light58 ' 2 player up
 set lights(59)=light59 ' 3 player up
 set lights(60)=light60 ' 4 player up
 set lights(61)=light61 ' tilt
 set lights(62)=light62 ' game over
 set lights(63)=light63 ' same player shoots again in backbox
 set lights(64)=light64 ' hight score


' Switches

Sub Sw18_hit():vpmtimer.pulsesw 18:End Sub
Sub Sw20_hit():vpmtimer.pulsesw 20:End Sub
Sub Sw21_hit():vpmtimer.pulsesw 21:End Sub
Sub Sw23_hit():vpmtimer.pulsesw 23:End Sub
Sub Sw30_hit():vpmtimer.pulsesw 30:End Sub
Sub Sw34_hit():vpmtimer.pulsesw 34:End Sub
Sub Sw43_hit():vpmtimer.pulsesw 43:End Sub
Sub Sw45_hit():vpmtimer.pulsesw 45:End Sub
Sub Sw46_hit():vpmtimer.pulsesw 46:End Sub
Sub Sw52_hit():vpmtimer.pulsesw 52:End Sub
Sub Sw53_hit():vpmtimer.pulsesw 53:End Sub


' Captive balls

 Sub CircularTarget1_hit()
  vpmtimer.pulsesw 31
  CapturedBallPost1.isdropped=0
  CircularTarget1.IsDropped = True
  CircularTarget1.TimerEnabled = True
 End Sub

 Sub CircularTarget2_hit()
  vpmtimer.pulsesw 32
  CapturedBallPost2.isdropped=0
  CircularTarget2.IsDropped = True
  CircularTarget2.TimerEnabled = True
 End Sub


 Sub CircularTarget3_hit()
  vpmtimer.pulsesw 33
  CapturedBallPost3.isdropped=0
  CircularTarget3.IsDropped = True
  CircularTarget3.TimerEnabled = True
 End Sub

Sub CircularTarget1_Timer()
  CircularTarget1.IsDropped = False
  CircularTarget1.TimerEnabled = False
End Sub

Sub CircularTarget2_Timer()
  CircularTarget2.IsDropped = False
  CircularTarget2.TimerEnabled = False
End Sub

Sub CircularTarget3_Timer()
  CircularTarget3.IsDropped = False
  CircularTarget3.TimerEnabled = False
End Sub

 Sub LeftOutlane_Hit()
  controller.switch (13)=true
  Switch2.isDropped= False
 End Sub

' Triggers

Sub sw11_Hit:Controller.Switch(11) = 1:End Sub
Sub sw11_UnHit:Controller.Switch(11) = 0:End Sub

Sub sw13_Hit:Controller.Switch(13) = 1:End Sub
Sub sw13_UnHit:Controller.Switch(13) = 0:End Sub

Sub sw12_Hit:Controller.Switch(12) = 1:End Sub
Sub sw12_UnHit:Controller.Switch(12) = 0:End Sub

Sub sw24_Hit:Controller.Switch(24) = 1:End Sub
Sub sw24_UnHit:Controller.Switch(24) = 0:End Sub

Sub sw25_Hit:Controller.Switch(25) = 1:End Sub
Sub sw25_UnHit:Controller.Switch(25) = 0:End Sub

Sub sw27_Hit:Controller.Switch(27) = 1:End Sub
Sub sw27_UnHit:Controller.Switch(27) = 0:End Sub

Sub sw28_Hit:Controller.Switch(28) = 1:End Sub
Sub sw28_UnHit:Controller.Switch(28) = 0:End Sub

Sub sw35_Hit:Controller.Switch(35) = 1:End Sub
Sub sw35_UnHit:Controller.Switch(35) = 0:End Sub

Sub sw37_Hit:Controller.Switch(37) = 1:End Sub
Sub sw37_UnHit:Controller.Switch(37) = 0:End Sub

Sub sw44_Hit:Controller.Switch(44) = 1:End Sub
Sub sw44_UnHit:Controller.Switch(44) = 0:End Sub

Sub Kicker1_Hit():bshig.addball 0:Kicker1.TimerEnabled = True:Pkickarm.rotz=15: End Sub

Sub Kicker1_Timer()
  Kicker1.TimerEnabled = False
  Pkickarm.rotz=0
End Sub

Sub Kicker2_Hit():bslow.addball 0:End Sub

Sub Spinner1_Spin():vpmtimer.pulsesw 22:End Sub
Sub Spinner2_Spin():vpmtimer.pulsesw 29:End Sub

Sub DropTarget1_Hit():dtbank3.hit 1:End sub
Sub DropTarget2_Hit():dtbank3.hit 2:End sub
Sub DropTarget3_Hit():dtbank3.hit 3:End sub

Sub DropTarget4_Hit():dtbank5.hit 1:End sub
Sub DropTarget5_Hit():dtbank5.hit 2:End sub
Sub DropTarget6_Hit():dtbank5.hit 3:End sub

Sub Gate1_Hit():PlaySoundAt "gate",Gate1:End Sub
Sub Gate2_Hit():PlaySoundAt "gate",Gate2:End Sub
Sub Gate3_Hit():PlaySoundAt "gate",Gate3:End Sub
Sub Gate4_Hit():PlaySoundAt "gate",Gate4:End Sub


'***** GI Lights On
dim xx

For each xx in GI:xx.State = 1: Next

If LMPFEnabled = 1 Then
	lightmap_pf.visible=1
Else
	lightmap_pf.visible=0
End If


'********** Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep

Sub RightSlingShot_Slingshot
    PlaySound SoundFX("right_slingshot",DOFContactors), 0,1, 0.05,0.05 '0,1, AudioPan(RightSlingShot), 0.05,0,0,1,AudioFade(RightSlingShot)
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.TransZ = -20
    RStep = 0
    RightSlingShot.TimerEnabled = 1
    vpmtimer.pulsesw 38
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.TransZ = -10
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.TransZ = 0:RightSlingShot.TimerEnabled = 0:
    End Select
    RStep = RStep + 1
End Sub

Sub URightSlingShot_Slingshot
    PlaySound SoundFX("right_slingshot",DOFContactors), 0,1, 0.05,0.05 '0,1, AudioPan(URightSlingShot), 0.05,0,0,1,AudioFade(URightSlingShot)
    URSling.Visible = 0
    URSling1.Visible = 1
    sling3.TransZ = -20
    RStep = 0
    URightSlingShot.TimerEnabled = 1
    vpmtimer.pulsesw 36
End Sub

Sub URightSlingShot_Timer
    Select Case RStep
        Case 3:URSLing1.Visible = 0:URSLing2.Visible = 1:sling3.TransZ = -10
        Case 4:URSLing2.Visible = 0:URSLing.Visible = 1:sling3.TransZ = 0:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
    PlaySound SoundFX("left_slingshot",DOFContactors), 0,1, -0.05,0.05 '0,1, AudioPan(LeftSlingShot), 0.05,0,0,1,AudioFade(LeftSlingShot)
    LSling.Visible = 0
    LSling1.Visible = 1
    Sling2.TransZ = -20
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
    vpmtimer.pulsesw 10
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.TransZ = -10
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.TransZ = 0:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub


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


'*****************************************
'   rothbauerw's Manual Ball Control
'*****************************************

Dim BCup, BCdown, BCleft, BCright
Dim ControlBallInPlay, ControlActiveBall
Dim BCvel, BCyveloffset, BCboostmulti, BCboost

BCboost = 1       'Do Not Change - default setting
BCvel = 4       'Controls the speed of the ball movement
BCyveloffset = -0.01  'Offsets the force of gravity to keep the ball from drifting vertically on the table, should be negative
BCboostmulti = 3    'Boost multiplier to ball veloctiy (toggled with the B key)

ControlBallInPlay = false

Sub StartBallControl_Hit()
  Set ControlActiveBall = ActiveBall
  ControlBallInPlay = true
End Sub

Sub StopBallControl_Hit()
  ControlBallInPlay = false
End Sub

Sub BallControlTimer_Timer()
  If EnableBallControl and ControlBallInPlay then
    If BCright = 1 Then
      ControlActiveBall.velx =  BCvel*BCboost
    ElseIf BCleft = 1 Then
      ControlActiveBall.velx = -BCvel*BCboost
    Else
      ControlActiveBall.velx = 0
    End If

    If BCup = 1 Then
      ControlActiveBall.vely = -BCvel*BCboost
    ElseIf BCdown = 1 Then
      ControlActiveBall.vely =  BCvel*BCboost
    Else
      ControlActiveBall.vely = bcyveloffset
    End If
  End If
End Sub


'*****************************************
' ninuzzu's FLIPPER SHADOWS
'*****************************************

'sub FlipperTimer_Timer()
'	if FlipperShadows = 1 then
'		FlipperLSh.RotZ = batleft.objrotz
'		FlipperRSh.RotZ = batright.objrotz
'  		FlipperLSh.RotZ = LeftFlipper.currentangle
'  		FlipperRSh.RotZ = RightFlipper.currentangle
'	end if
'End Sub

'*****************************************
' ninuzzu's BALL SHADOW
'*****************************************

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
            BallShadow(b).X = ((BOT(b).X) - (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/7)) '+ 6
        Else
            BallShadow(b).X = ((BOT(b).X) + (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/7)) '- 6
        End If
        ballShadow(b).Y = BOT(b).Y + 10
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

'**********************************************************************************************************
'Digital Display
'**********************************************************************************************************

Dim Digits(32)
' 1st Player
Digits(0) = Array(a1,a2,a3,a4,a5,a6,a7)
Digits(1) = Array(a8,a9,a10,a11,a12,a13,a14)
Digits(2) = Array(a15,a16,a17,a18,a19,a20,a21)
Digits(3) = Array(a22,a23,a24,a25,a26,a27,a28)
Digits(4) = Array(a29,a30,a31,a32,a33,a34,a35)
Digits(5) = Array(a36,a37,a38,a39,a40,a41,a42)
Digits(6) = Array(a43,a44,a45,a46,a47,a48,a49)

' 2nd Player
Digits(7) = Array(a50,a51,a52,a53,a54,a55,a56)
Digits(8) = Array(a57,a58,a59,a60,a61,a62,a63)
Digits(9) = Array(a64,a65,a66,a67,a68,a69,a70)
Digits(10) = Array(a71,a72,a73,a74,a75,a76,a77)
Digits(11) = Array(a78,a79,a80,a81,a82,a83,a84)
Digits(12) = Array(a85,a86,a87,a88,a89,a90,a91)
Digits(13) = Array(a92,a93,a94,a95,a96,a97,a98)

' 3rd Player
Digits(14) = Array(a99,a100,a101,a102,a103,a104,a105)
Digits(15) = Array(a106,a107,a108,a109,a110,a111,a112)
Digits(16) = Array(a113,a114,a115,a116,a117,a118,a119)
Digits(17) = Array(a120,a121,a122,a123,a124,a125,a126)
Digits(18) = Array(a127,a128,a129,a130,a131,a132,a133)
Digits(19) = Array(a134,a135,a136,a137,a138,a139,a140)
Digits(20) = Array(a141,a142,a143,a144,a145,a146,a147)

' 4th Player
Digits(21) = Array(a148,a149,a150,a151,a152,a153,a154)
Digits(22) = Array(a155,a156,a157,a158,a159,a160,a161)
Digits(23) = Array(a162,a163,a164,a165,a166,a167,a168)
Digits(24) = Array(a169,a170,a171,a172,a173,a174,a175)
Digits(25) = Array(a176,a177,a178,a179,a180,a181,a182)
Digits(26) = Array(a183,a184,a185,a186,a187,a188,a189)
Digits(27) = Array(a190,a191,a192,a193,a194,a195,a196)

' Credits
Digits(28) = Array(a197,a198,a199,a200,a201,a202,a203)
Digits(29) = Array(a204,a205,a206,a207,a208,a209,a210)
' Balls
Digits(30) = Array(a211,a212,a213,a214,a215,a216,a217)
Digits(31) = Array(a218,a219,a220,a221,a222,a223,a224)

Sub DisplayTimer_Timer
	Dim ChgLED,ii,num,chg,stat,obj
	ChgLed = Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
If Not IsEmpty(ChgLED) Then
		If DesktopMode = True Then
		For ii = 0 To UBound(chgLED)
			num = chgLED(ii, 0) : chg = chgLED(ii, 1) : stat = chgLED(ii, 2)
			if (num < 32) then
				For Each obj In Digits(num)
					If chg And 1 Then obj.State = stat And 1
					chg = chg\2 : stat = stat\2
				Next
			else
			end if
		next
		end if
end if
End Sub


 '=========================================================
'                    LED Handling
'=========================================================
'Modified version of Scapino's LED code for Fathom
'
'Dim SevenDigitOutput(32)
'Dim DisplayPatterns(11)
'Dim DigStorage(32)
'
'dim ledstatus : ledstatus = 2
'
''Binary/Hex Pattern Recognition Array
'DisplayPatterns(0) = 0   '0000000 Blank
'DisplayPatterns(1) = 63    '0111111 zero
'DisplayPatterns(2) = 6   '0000110 one
'DisplayPatterns(3) = 91    '1011011 two
'DisplayPatterns(4) = 79    '1001111 three
'DisplayPatterns(5) = 102 '1100110 four
'DisplayPatterns(6) = 109 '1101101 five
'DisplayPatterns(7) = 125 '1111101 six
'DisplayPatterns(8) = 7   '0000111 seven
'DisplayPatterns(9) = 127 '1111111 eight
'DisplayPatterns(10)= 111 '1101111 nine
'
''Assign 7-digit output to reels
'Set SevenDigitOutput(0)  = P3D7
'Set SevenDigitOutput(1)  = P3D6
'Set SevenDigitOutput(2)  = P3D5
'Set SevenDigitOutput(3)  = P3D4
'Set SevenDigitOutput(4)  = P3D3
'Set SevenDigitOutput(5)  = P3D2
'Set SevenDigitOutput(6)  = P3D1
'
'Set SevenDigitOutput(7)  = P4D7
'Set SevenDigitOutput(8)  = P4D6
'Set SevenDigitOutput(9)  = P4D5
'Set SevenDigitOutput(10) = P4D4
'Set SevenDigitOutput(11) = P4D3
'Set SevenDigitOutput(12) = P4D2
'Set SevenDigitOutput(13) = P4D1
'
'Set SevenDigitOutput(14) = P1D7
'Set SevenDigitOutput(15) = P1D6
'Set SevenDigitOutput(16) = P1D5
'Set SevenDigitOutput(17) = P1D4
'Set SevenDigitOutput(18) = P1D3
'Set SevenDigitOutput(19) = P1D2
'Set SevenDigitOutput(20) = P1D1
'
'Set SevenDigitOutput(21) = P2D7
'Set SevenDigitOutput(22) = P2D6
'Set SevenDigitOutput(23) = P2D5
'Set SevenDigitOutput(24) = P2D4
'Set SevenDigitOutput(25) = P2D3
'Set SevenDigitOutput(26) = P2D2
'Set SevenDigitOutput(27) = P2D1
'
'Set SevenDigitOutput(28) = CrD2
'Set SevenDigitOutput(29) = CrD1
'Set SevenDigitOutput(30) = BaD2
'Set SevenDigitOutput(31) = BaD1
'
'Sub DisplayTimer7_Timer ' 7-Digit output
' On Error Resume Next
' Dim ChgLED,ii,chg,stat,obj,TempCount,temptext,adj
'
' ChgLED = Controller.ChangedLEDs(&HFF, &HFFFF) 'hex of binary (display 111111, or first 6 digits)
'
' If Not IsEmpty(ChgLED) Then
'   For ii = 0 To UBound(ChgLED)
'     chg = chgLED(ii, 1) : stat = chgLED(ii, 2)
'     For TempCount = 0 to 10
'       If stat = DisplayPatterns(TempCount) then
'         If LedStatus = 2 Then SevenDigitOutput(chgLED(ii, 0)).SetValue(TempCount)
'         DigStorage(chgLED(ii, 0)) = TempCount
'       End If
'       If stat = (DisplayPatterns(TempCount) + 128) then
'         If LedStatus = 2 Then SevenDigitOutput(chgLED(ii, 0)).SetValue(TempCount)
'         DigStorage(chgLED(ii, 0)) = TempCount
'       End If
'     Next
'   Next
' End IF
'End Sub
'

' Williams Flippers

Sub GraphicsTimer_Timer()
	batleft.objrotz = LeftFlipper.CurrentAngle + 1
	DiverterP.objrotz = Outlanegate.CurrentAngle + 90
	batright.objrotz = RightFlipper.CurrentAngle - 1
	if FlipperShadows = True then
		FlipperLSh.RotZ = batleft.objrotz
		FlipperRSh.RotZ = batright.objrotz
	end if
End Sub

' Thalamus : Exit in a clean and proper way
Sub Table1_exit()
  Controller.Pause = False
  Controller.Stop
End Sub

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


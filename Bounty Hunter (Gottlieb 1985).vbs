Option Explicit

Const cGameName="bountyh",UseSolenoids=2,UseLamps=1,UseGI=0,SSolenoidOn="SolOn",SSolenoidOff="SolOff",SFlipperOn="FlipperDown",SFlipperOff="FlipperUp"
Const SCoin="coin3",cCredits=""

' Thalamus 2019 March : Improved directional sounds
' !! NOTE : Table not verified yet !!
' Table didn't have ballrolling at all, code section imported, need to add timer named RollingTimer value of 40.
' Also, add  rollover'
'

' Options
' Volume devided by - lower gets higher sound

Const VolDiv = 2000    ' Lower number, louder ballrolling/collition sound
Const VolCol = 10      ' Ball collition divider ( voldiv/volcol )

' The rest of the values are multipliers
'
'  .5 = lower volume
' 1.5 = higher volume

Const VolRol    = 1    ' Rollovers volume.
Const VolFlip   = 1    ' Flipper volume.

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

LoadVPM "01120100","sys80.vbs",3.02

Dim VarRol
If Table1.ShowDT = true then VarRol=0 Else VarRol=1

Set LampCallback=GetRef("UpdateMultipleLamps")

Sub UpdateMultipleLamps
End Sub

Sub Table1_KeyDown(ByVal keycode)
	If KeyCode=LeftFlipperKey Then Controller.Switch(6)=1
	If keyCode=RightFlipperKey Then Controller.Switch(16)=1
	If vpmKeyDown(KeyCode) Then Exit Sub
	If keycode=PlungerKey Then Plunger.Pullback
End Sub

Sub Table1_KeyUp(ByVal keycode)
	If KeyCode=LeftFlipperKey Then Controller.Switch(6)=0
	If keyCode=RightFlipperKey Then Controller.Switch(16)=0
	If vpmKeyUp(KeyCode) Then Exit Sub
	If keycode=PlungerKey Then Plunger.Fire:playsoundAtVol "plunger", plunger, 1
End Sub

Const sDropDown=1
Const sKicker=2
Const sDropUp=5
Const sknocker=8
Const sOutHole=9

SolCallback(sKicker)="bsSaucer.SolOut"
SolCallback(sKnocker)="VpmSolSound SoundFX(""knocker"",DOFKnocker),"
SolCallback(sDropDown)="SolDown"
SolCallback(sDropUp)="SolDro"
SolCallback(sOutHole)="bsTrough.SolOut"
'SolCallback(sLLFlipper)="VpmSolFlipper LeftFlipper,nothing,"
'SolCallback(sLRFlipper)="VpmSolFlipper RightFlipper,Flipper1,"

 SolCallback(sLRFlipper) = "SolRFlipper"
 SolCallback(sLLFlipper) = "SolLFlipper"

	Sub SolLFlipper(Enabled)
		If Enabled Then
		PlaySoundAtVol SoundFX("flipperup",DOFFlippers), LeftFlipper, VolFlip
		LeftFlipper.RotateToEnd
     Else
		 PlaySoundAtVol SoundFX("flipperdown",DOFFlippers), LeftFlipper, VolFlip
		 LeftFlipper.RotateToStart
     End If
	End Sub

	Sub SolRFlipper(Enabled)
     If Enabled Then
		 PlaySoundAtVol SoundFX("flipperup",DOFFlippers), RightFlipper, VolFlip
     PlaySoundAtVol "flipperup", Flipper1, VolFlip
		 RightFlipper.RotateToEnd
		 Flipper1.RotateToEnd
     Else
		 PlaySoundAtVol SoundFX("flipperdown",DOFFlippers), RightFlipper, VolFlip
     PlaySoundAtVol "flipperdown", Flipper1, VolFlip
		 RightFlipper.RotateToStart
		 Flipper1.RotateToStart
     End If
	End Sub

Sub SolDro(Enabled)
	If Enabled Then
		Wall4.IsDropped=0
		Controller.Switch(74)=0
	End If
End Sub

Sub SolDown(Enabled)
	If Enabled Then
		Wall4.IsDropped=1
		Controller.Switch(74)=1
	End If
End Sub

Dim bsTrough,bsSaucer

Sub Table1_Init
	On Error Resume Next
	With Controller
		.GameName=cGameName
		If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
		.SplashInfoLine="Bounty Hunter, Gottlieb/Premier 1985."
		.HandleMechanics=0
		.HandleKeyboard=0
		.ShowDMDOnly=1
		.ShowFrame=1
		.Games(cGameName).Settings.Value("rol") = VarRol
		.ShowTitle=0
		If Err Then MsgBox Err.Description
	End With
	On Error Goto 0
		Controller.SolMask(0)=0
      vpmTimer.AddTimer 2000,"Controller.SolMask(0)=&Hffffffff'" 'ignore all solenoids - then add the timer to renable all the solenoids after 2 seconds
		Controller.Run


'	Controller.Dip(0) = (0*1 + 0*2 + 0*4 + 0*8 + 0*16 + 0*32 + 0*64 + 1*128) '01-08
'	Controller.Dip(1) = (0*1 + 0*2 + 0*4 + 0*8 + 0*16 + 0*32 + 1*64 + 1*128) '09-16
'	Controller.Dip(2) = (0*1 + 0*2 + 0*4 + 0*8 + 0*16 + 1*32 + 1*64 + 1*128) '17-24
'	Controller.Dip(3) = (1*1 + 1*2 + 0*4 + 0*8 + 1*16 + 0*32 + 0*64 + 0*128) '25-32

'Switch 25 Number Of Balls
'    ON = 3
'   OFF = 5

	PinMAMETimer.Interval=PinMAMEInterval
	PinMAMETimer.Enabled=1

	vpmNudge.TiltSwitch=57
	vpmNudge.Sensitivity=1
	vpmNudge.TiltObj=Array(Bumper1,Bumper2,LeftSlingshot,RightSlingshot,Wall5,Wall6)

	Set bsTrough=New cvpmBallstack
	bsTrough.InitSw 0,67,0,0,0,0,0,0
	bsTrough.InitKick BallRelease,90,6
	bsTrough.InitExitSnd SoundFX("ballrel",DOFContactors),SoundFX("solon",DOFContactors)
	bsTrough.Balls=1

	Set bsSaucer=New cvpmBallStack
	bsSaucer.InitSaucer Kicker1,73,147,5
	bsSaucer.InitExitSnd SoundFX("popper",DOFContactors),SoundFX("popper",DOFContactors)

	Wall4.IsDropped=1
	Controller.Switch(74)=1
End Sub

Sub Trigger4_Hit:Controller.Switch(40)=1:End Sub				'switch 40
Sub Trigger4_unHit:Controller.switch(40)=0:End Sub				'switch 40
Sub Trigger5_Hit:Controller.Switch(41)=1:End Sub				'switch 41
Sub Trigger5_unHit:Controller.switch(41)=0:End Sub				'switch 41
Sub Trigger6_Hit:Controller.Switch(42)=1:End Sub				'switch 42
Sub Trigger6_unHit:Controller.switch(42)=0:End Sub				'switch 42
Sub Bumper1_Hit()PlaysoundAtVol SoundFX("bump",DOFContactors) , ActiveBall, 1:vpmTimer.PulseSw(43):Light1.State=1:Light1.TimerEnabled=True:DOF 103, DOFPulse:End Sub					'switch 43
Sub Light1_Timer:Light1.TimerEnabled=False:Light1.State=0:End Sub
Sub Bumper2_Hit()PlaysoundAtVol SoundFX("bump",DOFContactors), ActiveBall, 1 :vpmTimer.PulseSw(43):Light2.State=1:Light2.TimerEnabled=True:DOF 104, DOFPulse:End Sub					'switch 43
Sub Light2_Timer:Light2.TimerEnabled=False:Light2.State=0:End Sub
Sub Wall5_Slingshot:PlaySoundAtBallVol SoundFX("sling",DOFContactors),1:VpmTimer.PulseSw 44:DOF 106, DOFPulse:End Sub			'switch 44
Sub Wall6_Slingshot:PlaySoundAtBallVol SoundFX("sling",DOFContactors),1:VpmTimer.PulseSw 44:DOF 107, DOFPulse:End Sub			'switch 44
Sub LeftSlingShot_Slingshot:PlaySoundAtVol SoundFX("sling",DOFContactors),LeftSlingShot,1:VpmTimer.PulseSw 44:DOF 101, DOFPulse:End Sub 'switch 44
Sub RightSlingshot_Slingshot:PlaySoundAtVol SoundFX("sling",DOFContactors),RightSlingShot,1:VpmTimer.PulseSw 44:DOF 102, DOFPulse:End Sub 'switch 44
Sub Wall1_Hit:VpmTimer.PulseSw 50:End Sub						'switch 50
Sub Wall7_Hit:vpmTimer.PulseSw 51:End Sub						'switch 51
Sub Trigger1_Hit:Controller.Switch(52)=1:End Sub				'switch 52
Sub Trigger1_unHit:Controller.switch(52)=0:End Sub				'switch 52
Sub RightInlane_Hit:Controller.Switch(53)=1:End Sub				'switch 53
Sub RightInlane_UnHit:Controller.Switch(53)=0:End Sub 			'switch 53
Sub LeftInlane_Hit:Controller.Switch(54)=1:End Sub				'switch 54
Sub LeftInlane_unHit:Controller.Switch(54)=0:End Sub			'switch 54
Sub Wall2_Hit:VpmTimer.PulseSw 60:End Sub						'switch 60
Sub Wall8_Hit:vpmTimer.PulseSw 61:End Sub						'switch 61
Sub Trigger2_Hit:Controller.Switch(62)=1:End Sub				'switch 62
Sub Trigger2_unHit:Controller.switch(62)=0:End Sub				'switch 62
Sub RightOutlane_Hit:Controller.Switch(63)=1:End Sub			'switch 63
Sub RightOutlane_UnHit:Controller.Switch(63)=0:End Sub			'switch 63
Sub LeftOutlane_Hit:Controller.Switch(64)=1:End Sub				'switch 64
Sub LetOutlane_unHit:Controller.Switch(64)=0:End Sub			'switch 64
Sub Drain_Hit()PlaysoundAtVol "drain", drain, 1:bsTrough.AddBall Me:End Sub						'switch 67
Sub Wall3_Hit:VpmTimer.PulseSw 70:End Sub						'switch 70
Sub Wall9_Hit:vpmTimer.PulseSw 71:End Sub						'switch 71
Sub Trigger3_Hit:Controller.Switch(72)=1:End Sub				'switch 72
Sub Trigger3_unHit:Controller.switch(72)=0:End Sub				'switch 72
Sub Kicker1_Hit:bsSaucer.AddBall 0:End Sub						'switch 73
Sub Wall4_Hit:Wall4.IsDropped=1:Controller.Switch(74)=1:PlaySoundAtBallVol"flapclos",1:End Sub'switch 74

set lights(3)=Light3
Set Lights(6)=Light6
Set Lights(7)=Light7
Set Lights(8)=Light8
Set Lights(9)=Light9
Set Lights(10)=Light10
Set lights(11)=light11
set lights(12)=light12
lights(13)=Array(light13,light13B)
set lights(14)=light14
set lights(15)=light15
set lights(16)=light16
lights(17)=Array(light17,light17B)
set lights(18)=light18
set lights(19)=light19
set lights(20)=light20
set lights(21)=light21
set lights(22)=light22
set lights(23)=light23
lights(24)=Array(light24,light24B)
set lights(25)=light25
lights(26)=Array(light26,light26B)
set lights(27)=light27
set lights(28)=light28
set lights(29)=light29
set lights(30)=light30
set lights(31)=light31
set lights(32)=light32
set lights(33)=light33
set lights(34)=light34
set lights(35)=light35
set lights(36)=light36
set lights(37)=light37
set lights(38)=light38
set lights(39)=light39
set lights(40)=light40
set lights(41)=light41
set lights(42)=light42
set lights(43)=light43
set lights(44)=light44
set lights(45)=light45
set lights(46)=light46
lights(47)=Array(light47, light48, light49, light50, light51, light52, light53, light54, light55, light56, light57, light58, light59, light60, light61, light62, light63)

Sub Gate1_Hit: PlaySoundAtVol "Gate" , ActiveBall, 1: End Sub
Sub Gate_Hit: PlaySoundAtVol "Gate" , ActiveBall, 1: End Sub
Sub Trigger7_Hit: PlaySoundAtVol "PlungeBall" , ActiveBall, 1: End Sub

'Gottlieb System 80A Sound only (S)board
'added by Inkochnito
Sub editDips
	Dim vpmDips : Set vpmDips = New cvpmDips
	With vpmDips
		.AddForm 700,400,"System 80A with Sound only (S)board - DIP switches"
		.AddFrame 0,0,190,"Maximum credits",49152,Array("8 credits",0,"10 credits",32768,"15 credits",&H00004000,"25 credits",49152)'dip 15&16
		.AddFrame 0,76,190,"Coin chute 1 and 2 control",&H00002000,Array("seperate",0,"same",&H00002000)'dip 14
		.AddFrame 0,122,190,"3rd coin chute credits control",&H20000000,Array("no effect",0,"add 9",&H20000000)'dip 30
		.AddFrame 205,0,190,"High score to date awards",&H00C00000,Array("not displayed and no award",0,"displayed and no award",&H00800000,"displayed and 2 credits",&H00400000,"displayed and 3 credits",&H00C00000)'dip 23&24
		.AddFrame 205,76,190,"Game mode",&H10000000,Array("replay",0,"extra ball",&H10000000)'dip 29
		.AddFrame 205,122,190,"Playfield special",&H00200000,Array("replay",0,"extra ball",&H00200000)'dip 22
		.AddFrame 205,168,190,"Replay limit",&H04000000,Array("no limit",0,"one per game",&H04000000)'dip 27
		.AddFrame 205,214,190,"Balls per game",&H01000000,Array("5 balls",0,"3 balls",&H01000000)'dip 25
		.AddFrame 205,260,190,"Novelty mode",&H08000000,Array("normal game mode",0,"50K per special/extra ball",&H08000000)'dip 28
		.AddChk 0,170,190,Array("Match feature",&H02000000)'dip 26
		.AddChk 0,185,190,Array("Background sound",&H40000000)'dip 31
		.AddChk 0,200,190,Array("Dip 6 (spare)",&H00000020)'dip 6
		.AddChk 0,215,190,Array("Dip 7 (spare)",&H00000040)'dip 7
		.AddChk 0,230,190,Array("Dip 8 (spare)",&H00000080)'dip 8
		.AddChk 0,245,190,Array("Dip 32 (spare or game option)",&H80000000)'dip 32
		.AddLabel 50,310,300,20,"After hitting OK, press F3 to reset game with new settings."
		.ViewDips
	End With
End Sub
Set vpmShowDips = GetRef("editDips")

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

Sub Switches_Hit (idx)
  PlaySound "rollover", 0, Vol(ActiveBall)*VolRol, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Randomize

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
Const VolRol    = 1    ' Rollovers volume.
Const VolGates  = 1    ' Gates volume.
Const VolMetal  = 1    ' Metals volume.
Const VolRB     = 1    ' Rubber bands volume.
Const VolRH     = 1    ' Rubber hits volume.
Const VolPo     = 1    ' Rubber posts volume.
Const VolPi     = 1    ' Rubber pins volume.
Const VolPlast  = 1    ' Plastics volume.
Const VolTarg   = 1    ' Targets volume.
Const VolWood   = 1    ' Woods volume.
Const VolKick   = 1    ' Kicker volume.
Const VolSpin   = 1.5  ' Spinners volume.
Const VolFlip   = 1    ' Flipper volume.


On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Const cGameName = "heat_wave_1964"

Dim Score
Dim BallstoPlay
Dim Credits
Dim GameinProgress
Dim GameStart
Dim BallsPerGame
Dim Trumpet
Dim Randy
Dim Advancer
Dim Nathan
Dim Dathan
Dim Candy
Dim Vidi
Dim Freeze
Dim Gee
Dim Over
Dim Postman
Dim Shep
Dim Omar
Dim Highscore
Dim DesAngel
dim TargPos
dim TargDir
Dim TempLevel
Dim Tilt
Dim TiltSensor
Dim BIP
Dim Replay
Dim BallinLane
Dim Playedballs
Dim timer1flag,timer2flag,timer3flag,timer4flag,timer5flag,timer6flag
'-----------------------------------------------------
Dim Matchnum
'Not sure if this is necessary------------------------

 ExecuteGlobal GetTextFile("core.vbs")

Dim DesktopMode: DesktopMode = Heatw.ShowDT

If DesktopMode = True Then
	EMReel1.Visible=True
	HighScoreText.visible=true
	Sbest.visible=true
	t1.visible=true
	t2.visible=true
	t3.visible=true
	t4.visible=true
	t5.visible=true
	t6.visible=true
	t7.visible=true
	t8.visible=true
	t9.visible=true
	t10.visible=true
	t11.visible=true
	t12.visible=true
	t13.visible=true
	t14.visible=true
	t15.visible=true
	t16.visible=true
	t17.visible=true
	t18.visible=true
	t19.visible=true
	t20.visible=true
	t21.visible=true
	t22.visible=true
	t23.visible=true
	t24.visible=true
	t25.visible=true
	top.visible=true
	match0.visible=true
	match1.visible=true
	match2.visible=true
	match3.visible=true
	match4.visible=true
	match5.visible=true
	match6.visible=true
	match7.visible=true
	match8.visible=true
	match9.visible=true
	EMReelFrigid.visible=True
	EMReelCool.visible=True
	EMReelWarm.visible=True
	EMReelHot.visible=True
	EMReelTorrid.visible=True
	EMReelBlow.visible=True
	EMReelWilliams.visible=True
	GameOver.visible=True
	TiltLight.visible=True
	credittext.visible=True
	EMReelSun.visible=True
	Letter1.visible=True
	Letter2.visible=True
	Letter3.visible=True
	Letter4.visible=True
	Letter5.visible=True
	Letter6.visible=True
	Letter7.visible=True
	Letter8.visible=True
Else
	EMReel1.Visible=False
	HighScoreText.visible=false
	Sbest.visible=False
	t1.visible=false
	t2.visible=false
	t3.visible=false
	t4.visible=false
	t5.visible=false
	t6.visible=false
	t7.visible=false
	t8.visible=false
	t9.visible=false
	t10.visible=false
	t11.visible=false
	t12.visible=false
	t13.visible=false
	t14.visible=false
	t15.visible=false
	t16.visible=false
	t17.visible=false
	t18.visible=false
	t19.visible=false
	t20.visible=false
	t21.visible=false
	t22.visible=false
	t23.visible=false
	t24.visible=false
	t25.visible=false
	top.visible=false
	match0.visible=false
	match1.visible=false
	match2.visible=false
	match3.visible=false
	match4.visible=false
	match5.visible=false
	match6.visible=false
	match7.visible=false
	match8.visible=false
	match9.visible=false
	EMReelFrigid.visible=false
	EMReelCool.visible=false
	EMReelWarm.visible=false
	EMReelHot.visible=false
	EMReelTorrid.visible=false
	EMReelBlow.visible=false
	EMReelWilliams.visible=false
	GameOver.visible=false
	TiltLight.visible=false
	Credittext.visible=false
	EMReelSun.visible=false
	Letter1.visible=false
	Letter2.visible=false
	Letter3.visible=false
	Letter4.visible=false
	Letter5.visible=false
	Letter6.visible=false
	Letter7.visible=false
	Letter8.visible=false
end If




Sub Heatw_Init()

LoadEM

	timer1flag=0
	timer2flag=0
	timer3flag=0
	timer4flag=0
	timer5flag=0
	timer6flag=0
timer1.enabled=true
timer2.enabled=true
timer3.enabled=true
timer4.enabled=true
timer5.enabled=true
timer6.enabled=true
ballinlane=1
    	If B2SOn then Controller.B2SSetGameOver 1
			If B2SOn then Controller.B2SSetScoreRolloverPlayer1 0
		If B2SOn then Controller.B2SSetScorePlayer1 0
		If B2SOn then Controller.B2SSetTilt 0
		If B2SOn then Controller.B2SSetCredits Credits
		For i = 51 to 77
			If B2SOn then Controller.B2SSetData i,0
		next
		If B2SOn then Controller.B2SSetData 51,1
		For i = 81 to 87
			If B2SOn then Controller.B2SSetData i,0
		next

LeftTarget2.IsDropped=1
LeftTarget3.IsDropped=1
LeftTarget4.IsDropped=1
LeftTarget5.IsDropped=1
LeftTarget6.IsDropped=1
LeftTarget7.IsDropped=1
LeftTarget8.IsDropped=1
LeftTarget9.IsDropped=1
LeftTarget10.IsDropped=1
LeftTarget11.IsDropped=1
LeftTarget12.IsDropped=1
LeftTarget13.IsDropped=1
LeftTarget14.IsDropped=1
LeftTarget15.IsDropped=1
LeftTarget16.IsDropped=1
LeftTarget17.IsDropped=1
LeftTarget18.IsDropped=1
LeftTarget19.IsDropped=1
LeftTarget20.IsDropped=1
LeftTarget21.IsDropped=1
LeftTarget22.IsDropped=1
LeftTarget23.IsDropped=1
LeftTarget24.IsDropped=1
LeftTarget25.IsDropped=1
LeftTarget26.IsDropped=1
LeftTarget27.IsDropped=1
LeftTarget28.IsDropped=1
LeftTarget29.IsDropped=1
LeftTarget30.IsDropped=1
LeftTarget31.IsDropped=1
LeftTarget32.IsDropped=1
LeftTarget33.IsDropped=1
LeftTarget34.IsDropped=1
LeftTarget35.IsDropped=1
LeftTarget36.IsDropped=1
LeftTarget37.IsDropped=1
LeftTarget38.IsDropped=1
LeftTarget39.IsDropped=1
LeftTarget40.IsDropped=1
LeftTarget41.IsDropped=1
LeftTarget42.IsDropped=1
LeftTarget43.IsDropped=1
LeftTarget44.IsDropped=1
LeftTarget45.IsDropped=1
LeftTarget46.IsDropped=1
LeftTarget47.IsDropped=1
LeftTarget48.IsDropped=1
LeftTarget49.IsDropped=1
LeftTarget50.IsDropped=1
rightTarget52.IsDropped=1
rightTarget53.IsDropped=1
rightTarget54.IsDropped=1
rightTarget55.IsDropped=1
rightTarget56.IsDropped=1
rightTarget57.IsDropped=1
rightTarget58.IsDropped=1
rightTarget59.IsDropped=1
rightTarget60.IsDropped=1
rightTarget61.IsDropped=1
rightTarget62.IsDropped=1
rightTarget63.IsDropped=1
rightTarget64.IsDropped=1
rightTarget65.IsDropped=1
rightTarget66.IsDropped=1
rightTarget67.IsDropped=1
rightTarget68.IsDropped=1
rightTarget69.IsDropped=1
rightTarget70.IsDropped=1
rightTarget71.IsDropped=1
rightTarget72.IsDropped=1
rightTarget73.IsDropped=1
rightTarget74.IsDropped=1
rightTarget75.IsDropped=1
rightTarget76.IsDropped=1
rightTarget77.IsDropped=1
rightTarget78.IsDropped=1
rightTarget79.IsDropped=1
rightTarget80.IsDropped=1
rightTarget81.IsDropped=1
rightTarget82.IsDropped=1
rightTarget83.IsDropped=1
rightTarget84.IsDropped=1
rightTarget85.IsDropped=1
rightTarget86.IsDropped=1
rightTarget87.IsDropped=1
rightTarget88.IsDropped=1
rightTarget89.IsDropped=1
rightTarget90.IsDropped=1
rightTarget91.IsDropped=1
rightTarget92.IsDropped=1
rightTarget93.IsDropped=1
rightTarget94.IsDropped=1
rightTarget95.IsDropped=1
rightTarget96.IsDropped=1
rightTarget97.IsDropped=1
rightTarget98.IsDropped=1
rightTarget99.IsDropped=1
rightTarget100.IsDropped=1
motor.enabled=false
Replay=0
BIP=0
TiltSensor=0
TempLevelLights.Enabled=true
TempLevel=1
TargPos=1
TargDir=1
TargetTimer.Enabled=false
HighScore=LoadValue("Heatw","HS") ' Load saved Highscore
	If HighScore="" then
		HighScore=100
		SaveValue "Heatw","HS",HighScore
	Else
		Highscore=CDbl(LoadValue("Heatw","HS"))
	End If
	Sbest.Text=FormatNumber(HighScore, 0, -1, 0, -1)
'----------------------------------------------------
Matchnum=0
'Set the variable-------------------------------------
Credits=0
GameinProgress=0
GameStart=0
	Score=0
   BallsPerGame=5
   BallstoPlay=0
   Credittext.Text=Credits
   Operator.text="Hit O to toggle operator mode between 3 and 5 ball play"
   ball.text=ballspergame
	If Credits > 0 Then DOF 121, DOFOn

If HighScore > 1000 then
	Instructions.visible=false
Else
	Instructions.visible=true
End If
 End Sub

Sub EndofGame()
Dim Score2
Dim Score3
Dim Score4
Dim MatchPick
Matchpick = int (rnd*9)
if Matchpick = 0 then
	If B2SOn then Controller.B2SSetMatch 10
else
	If B2SOn then Controller.B2SSetMatch Matchpick
end if
if Matchpick = 0 then Match0.SetValue 1 End If
if Matchpick = 1 then Match1.SetValue 1 End If
if Matchpick = 2 then Match2.SetValue 1 End If
if Matchpick = 3 then Match3.SetValue 1 End If
if Matchpick = 4 then Match4.SetValue 1 End If
if Matchpick = 5 then Match5.SetValue 1 End If
if Matchpick = 6 then Match6.SetValue 1 End If
if Matchpick = 7 then Match7.SetValue 1 End If
if Matchpick = 8 then Match8.SetValue 1 End If
if Matchpick = 9 then Match9.SetValue 1 End If
			Score2 = Score / 10
			Score3 = int(Score2)
			Score4 = Score - (10 * Score3)
If Matchpick=Score4 then
playsound SoundFXDOF("knock",119,DOFPulse,DOFKnocker)
DOF 120, DOFPulse
If Credits<9 then Credits=Credits+1:DOF 121, DOFOn:end if
credittext.text=credits
If B2SOn then Controller.B2SSetCredits Credits
End If

If score>HighScore Then
HighScore=score
SaveValue "Heatw","HS",HighScore
Sbest.Text=FormatNumber(HighScore, 0, -1, 0, -1)
End if
'--------------------------------------------------------

	'Matchtext.Text=FormatNumber(Matchnum,0,-1,0,-1)

'---------------------------------------------------------
GameinProgress=0
Shep=0
'Displays the number
	'If Matchnum = (Score Mod 10)Then
'Takes off last digit and checks it

		'Credittext.text=Credits
		'playsound "knock"
		'End If
End Sub


Sub Heatw_KeyDown(ByVal keycode)
'If GameStart=1 and GameinProgress=1 then
	If keycode = PlungerKey Then
		Plunger.PullBack
	End If

	If keycode = LeftFlipperKey Then
If Tilt=0 then
If GameInProgress=1 then
		LeftFlipper.RotateToEnd
		PlaySoundAtVol SoundFXDOF("FlipperUp",101,DOFOn,DOFContactors), LeftFlipper, VolFlip
End If
End If
	End If

	If keycode = RightFlipperKey Then
If Tilt=0 then
If GameinProgress=1 then
		RightFlipper.RotateToEnd
		PlaySoundAtVol SoundFXDOF("FlipperUp",102,DOFOn,DOFContactors), RightFlipper, VolFlip
End If
End If
	End If

    if keycode = LeftTiltKey Then
If BIP=1 Then
     Nudge 90, 2
TiltSensor=TiltSensor+50
End If
     End If

    If keycode = RightTiltKey Then
If BIP=1 Then
    Nudge 270, 2
TiltSensor=TiltSensor+50
End If
     End If

	If keycode = CenterTiltKey Then
If BIP=1 Then
		Nudge 0, 2
TiltSensor=TiltSensor+50
	End If

If keycode =  MechanicalTilt Then
  If BIP=1 Then
    TiltSensor=TiltSensor+50
	End If
End If


End If

if keycode = RightMagnaSave then
If Ballinlane=0 and playedballs<5 then
playsound "ballup"
BallRelease.isdropped=True
balllifter.enabled=true
End If
End If
if keycode = StartGameKey then
If Ballinlane=0 and playedballs<5 then
playsound "ballup"
BallRelease.isdropped=True
balllifter.enabled=true
End If
End If


	If Keycode = StartGameKey then
		If GameinProgress=0 and credits>0 then
playsound "motor"
motor.enabled=True
targettimer.enabled=true
ballinlane=0
playedballs=0
		Operator.text=""
		Ball.text=""
 		Credits=Credits-1
	Instructions.visible=false
		If Credits < 1 Then DOF 121, DOFOff
light1.state=0
 		Credittext.text=Credits
		If B2SOn then Controller.B2SSetCredits Credits
		If B2SOn then Controller.B2SSetGameOver 0 ' GameOver off
		If B2SOn then Controller.B2SSetPlayerUp 1
		If B2SOn then Controller.B2SSetMatch 0
		If B2SOn then Controller.B2SSetTilt 0
		If B2SOn then Controller.B2SSetScorePlayer1 0
		For i = 51 to 77
			If B2SOn then Controller.B2SSetData i,0
		next
		If B2SOn then Controller.B2SSetData 52,1
		For i = 81 to 87
			If B2SOn then Controller.B2SSetData i,0
		next
		If B2SOn then Controller.B2SSetData 81,1

 		GameStart=1
playsound "DrainShorter"
 		GameinProgress=1
Match0.SetValue 0
Match1.SetValue 0
Match2.SetValue 0
Match3.SetValue 0
Match4.SetValue 0
Match5.SetValue 0
Match6.SetValue 0
Match7.SetValue 0
Match8.SetValue 0
Match9.SetValue 0
GameOver.SetValue 0
If B2SOn then Controller.B2SSetGameOver 0
lightg1.state=1
lightg2.state=1
lightg3.state=1
lightg4.state=1
lightg5.state=1
Replay=0
    	Start_Game()
Wall49.IsDropped=1
   		End If
   	End IF


 	If Keycode= 6 then
 	playsoundAtVol "coin", drain, 1
  	If credits<9 then Credits=Credits+1:DOF 121, DOFOn:end if
  	Credittext.text=Credits
	If B2SOn then Controller.B2SSetCredits Credits
  	End if

End Sub

Sub Heatw_KeyUp(ByVal keycode)

    'If GameStart=1 and GameinProgress=1 then
	If keycode = PlungerKey Then
		Plunger.Fire
	PlaySoundAtVol "Plunger", plunger, 1
	End If


	If keycode = LeftFlipperKey Then
If Tilt=0 then
If GameInProgress=1 then
		LeftFlipper.RotateToStart
		PlaySoundAtVol SoundFXDOF("FlipperDown",101,DOFOff,DOFContactors), LeftFlipper, VolFlip
End If
End If
	End If

	If keycode = RightFlipperKey Then
If GameinProgress=1 then
If Tilt=0 then
		RightFlipper.RotateToStart

		PlaySoundAtVol SoundFXDOF("FlipperDown",102,DOFOff,DOFContactors), RightFlipper, VolFlip
End If
End If
	End If

	if keycode=24 and GameStart=0 and GameinProgress=0then
	call Toggleroutine()
	playsound "metal4"
	End If
	End Sub

	Sub Toggleroutine()
  if ballspergame=3 then
ballspergame=5
ball.text=ballspergame
else ballspergame=3
ball.text=ballspergame
end if
end sub

sub Motor_Timer
If GameStart=1 then playsound "motor"
End Sub

Sub BallLifter_Timer
	Kicker1.CreateBall.image="pinball"
	Kicker1.kick 270,2
ballinlane=1
playedballs=playedballs+1
BallLifter.enabled=false
BallLifter2.enabled=true
End Sub
Sub BallLifter2_Timer
	BallRelease.IsDropped=False
BallLifter2.enabled=false
End Sub

Sub TiltTimer_Timer
If TiltSensor>1 then TiltSensor=TiltSensor-1
If TiltSensor>99 then
TiltSensor=0
If Tilt=0 then playsound "buzzer" End If
Tilt=1
TiltLight.SetValue 1
If B2SOn then Controller.B2SSetTilt 1
End If
If light1.state=1 then emreelblow.setvalue 1 End If
If light1.state=0 then emreelblow.setvalue 0 End If
If sunlight.state=1 then emreelsun.setvalue 1 End If
If sunlight.state=0 then emreelsun.setvalue 0 End If
If williamslight.state=1 then emreelwilliams.setvalue 1 End If
If williamslight.state=0 then emreelwilliams.setvalue 0 End If
End Sub

Sub BlowTimer_timer
If light1.state=0 then
	light1.state=1
	If B2SOn then Controller.B2SSetData 86,1
else
	light1.state=0
	If B2SOn then Controller.B2SSetData 86,0
end if

End Sub

Sub SunTimer_timer
If sunlight.state=0 then sunlight.state=1 else sunlight.state=0
End Sub

Sub WilliamsTimer_timer
If williamslight.state=0 then williamslight.state=1 else williamslight.state=0
End Sub

Sub UpdateFlipperLogos_Timer
    LFLogo.ObjRotZ = LeftFlipper.CurrentAngle+80
    RFlogo.ObjRotZ = RightFlipper.CurrentAngle+80
	if maintarget.isdropped = 1 then shadows1.visible=false
	if maintarget.isdropped = 0 then shadows1.visible=true
End Sub

Sub TargetTimer_Timer
If TargPos=0 Then TargDir=1
If TargPos<6 Then LeftTarget1.IsDropped=0 else LeftTarget1.IsDropped=1 End If
If TargPos>5 and TargPos<10 Then LeftTarget2.IsDropped=0 else LeftTarget2.IsDropped=1 End If
If TargPos>9 and TargPos<13 Then LeftTarget3.IsDropped=0 else LeftTarget3.IsDropped=1 End If
If TargPos>12 and TargPos<15 Then LeftTarget4.IsDropped=0 else LeftTarget4.IsDropped=1 End If
If TargPos=15 Then LeftTarget5.IsDropped=0 else LeftTarget5.IsDropped=1 End If
If TargPos=16 Then LeftTarget6.IsDropped=0 else LeftTarget6.IsDropped=1 End If
If TargPos=17 Then LeftTarget7.IsDropped=0 else LeftTarget7.IsDropped=1 End If
If TargPos=18 Then LeftTarget8.IsDropped=0 else LeftTarget8.IsDropped=1 End If
If TargPos=19 Then LeftTarget9.IsDropped=0 else LeftTarget9.IsDropped=1 End If
If TargPos=20 Then LeftTarget10.IsDropped=0 else LeftTarget10.IsDropped=1 End If
If TargPos=21 Then LeftTarget11.IsDropped=0 else LeftTarget11.IsDropped=1 End If
If TargPos=22 Then LeftTarget12.IsDropped=0 else LeftTarget12.IsDropped=1 End If
If TargPos=23 Then LeftTarget13.IsDropped=0 else LeftTarget13.IsDropped=1 End If
If TargPos=24 Then LeftTarget14.IsDropped=0 else LeftTarget14.IsDropped=1 End If
If TargPos=25 Then LeftTarget15.IsDropped=0 else LeftTarget15.IsDropped=1 End If
If TargPos=26 Then LeftTarget16.IsDropped=0 else LeftTarget16.IsDropped=1 End If
If TargPos=27 Then LeftTarget17.IsDropped=0 else LeftTarget17.IsDropped=1 End If
If TargPos=28 Then LeftTarget18.IsDropped=0 else LeftTarget18.IsDropped=1 End If
If TargPos=29 Then LeftTarget19.IsDropped=0 else LeftTarget19.IsDropped=1 End If
If TargPos=30 Then LeftTarget20.IsDropped=0 else LeftTarget20.IsDropped=1 End If
If TargPos=31 Then LeftTarget21.IsDropped=0 else LeftTarget21.IsDropped=1 End If
If TargPos=32 Then LeftTarget22.IsDropped=0 else LeftTarget22.IsDropped=1 End If
If TargPos=33 Then LeftTarget23.IsDropped=0 else LeftTarget23.IsDropped=1 End If
If TargPos=34 Then LeftTarget24.IsDropped=0 else LeftTarget24.IsDropped=1 End If
If TargPos=35 Then LeftTarget25.IsDropped=0 else LeftTarget25.IsDropped=1 End If
If TargPos=36 Then LeftTarget26.IsDropped=0 else LeftTarget26.IsDropped=1 End If
If TargPos=37 Then LeftTarget27.IsDropped=0 else LeftTarget27.IsDropped=1 End If
If TargPos=38 Then LeftTarget28.IsDropped=0 else LeftTarget28.IsDropped=1 End If
If TargPos=39 Then LeftTarget29.IsDropped=0 else LeftTarget29.IsDropped=1 End If
If TargPos=40 Then LeftTarget30.IsDropped=0 else LeftTarget30.IsDropped=1 End If
If TargPos=41 Then LeftTarget31.IsDropped=0 else LeftTarget31.IsDropped=1 End If
If TargPos=42 Then LeftTarget32.IsDropped=0 else LeftTarget32.IsDropped=1 End If
If TargPos=43 Then LeftTarget33.IsDropped=0 else LeftTarget33.IsDropped=1 End If
If TargPos=44 Then LeftTarget34.IsDropped=0 else LeftTarget34.IsDropped=1 End If
If TargPos=45 Then LeftTarget35.IsDropped=0 else LeftTarget35.IsDropped=1 End If
If TargPos=46 Then LeftTarget36.IsDropped=0 else LeftTarget36.IsDropped=1 End If
If TargPos=47 Then LeftTarget37.IsDropped=0 else LeftTarget37.IsDropped=1 End If
If TargPos=48 Then LeftTarget38.IsDropped=0 else LeftTarget38.IsDropped=1 End If
If TargPos=49 Then LeftTarget39.IsDropped=0 else LeftTarget39.IsDropped=1 End If
If TargPos=50 Then LeftTarget40.IsDropped=0 else LeftTarget40.IsDropped=1 End If
If TargPos=51 Then LeftTarget41.IsDropped=0 else LeftTarget41.IsDropped=1 End If
If TargPos=52 Then LeftTarget42.IsDropped=0 else LeftTarget42.IsDropped=1 End If
If TargPos=53 Then LeftTarget43.IsDropped=0 else LeftTarget43.IsDropped=1 End If
If TargPos=54 Then LeftTarget44.IsDropped=0 else LeftTarget44.IsDropped=1 End If
If TargPos=55 Then LeftTarget45.IsDropped=0 else LeftTarget45.IsDropped=1 End If
If TargPos=56 Then LeftTarget46.IsDropped=0 else LeftTarget46.IsDropped=1 End If
If TargPos>56 and TargPos<59 Then LeftTarget47.IsDropped=0 else LeftTarget47.IsDropped=1 End If
If TargPos>58 and TargPos<62 Then LeftTarget48.IsDropped=0 else LeftTarget48.IsDropped=1 End If
If TargPos>61 and TargPos<65 Then LeftTarget49.IsDropped=0 else LeftTarget49.IsDropped=1 End If
If TargPos>64 Then LeftTarget50.IsDropped=0 else LeftTarget50.IsDropped=1 End If

If TargPos<6 Then RightTarget51.IsDropped=0 else RightTarget51.IsDropped=1 End If
If TargPos>5 and TargPos<10 Then RightTarget52.IsDropped=0 else RightTarget52.IsDropped=1 End If
If TargPos>9 and TargPos<13 Then RightTarget53.IsDropped=0 else RightTarget53.IsDropped=1 End If
If TargPos>12 and TargPos<15 Then RightTarget54.IsDropped=0 else RightTarget54.IsDropped=1 End If
If TargPos=15 Then RightTarget55.IsDropped=0 else RightTarget55.IsDropped=1 End If
If TargPos=16 Then RightTarget56.IsDropped=0 else RightTarget56.IsDropped=1 End If
If TargPos=17 Then RightTarget57.IsDropped=0 else RightTarget57.IsDropped=1 End If
If TargPos=18 Then RightTarget58.IsDropped=0 else RightTarget58.IsDropped=1 End If
If TargPos=19 Then RightTarget59.IsDropped=0 else RightTarget59.IsDropped=1 End If
If TargPos=20 Then RightTarget60.IsDropped=0 else RightTarget60.IsDropped=1 End If
If TargPos=21 Then RightTarget61.IsDropped=0 else RightTarget61.IsDropped=1 End If
If TargPos=22 Then RightTarget62.IsDropped=0 else RightTarget62.IsDropped=1 End If
If TargPos=23 Then RightTarget63.IsDropped=0 else RightTarget63.IsDropped=1 End If
If TargPos=24 Then RightTarget64.IsDropped=0 else RightTarget64.IsDropped=1 End If
If TargPos=25 Then RightTarget65.IsDropped=0 else RightTarget65.IsDropped=1 End If
If TargPos=26 Then RightTarget66.IsDropped=0 else RightTarget66.IsDropped=1 End If
If TargPos=27 Then RightTarget67.IsDropped=0 else RightTarget67.IsDropped=1 End If
If TargPos=28 Then RightTarget68.IsDropped=0 else RightTarget68.IsDropped=1 End If
If TargPos=29 Then RightTarget69.IsDropped=0 else RightTarget69.IsDropped=1 End If
If TargPos=30 Then RightTarget70.IsDropped=0 else RightTarget70.IsDropped=1 End If
If TargPos=31 Then RightTarget71.IsDropped=0 else RightTarget71.IsDropped=1 End If
If TargPos=32 Then RightTarget72.IsDropped=0 else RightTarget72.IsDropped=1 End If
If TargPos=33 Then RightTarget73.IsDropped=0 else RightTarget73.IsDropped=1 End If
If TargPos=34 Then RightTarget74.IsDropped=0 else RightTarget74.IsDropped=1 End If
If TargPos=35 Then RightTarget75.IsDropped=0 else RightTarget75.IsDropped=1 End If
If TargPos=36 Then RightTarget76.IsDropped=0 else RightTarget76.IsDropped=1 End If
If TargPos=37 Then RightTarget77.IsDropped=0 else RightTarget77.IsDropped=1 End If
If TargPos=38 Then RightTarget78.IsDropped=0 else RightTarget78.IsDropped=1 End If
If TargPos=39 Then RightTarget79.IsDropped=0 else RightTarget79.IsDropped=1 End If
If TargPos=40 Then RightTarget80.IsDropped=0 else RightTarget80.IsDropped=1 End If
If TargPos=41 Then RightTarget81.IsDropped=0 else RightTarget81.IsDropped=1 End If
If TargPos=42 Then RightTarget82.IsDropped=0 else RightTarget82.IsDropped=1 End If
If TargPos=43 Then RightTarget83.IsDropped=0 else RightTarget83.IsDropped=1 End If
If TargPos=44 Then RightTarget84.IsDropped=0 else RightTarget84.IsDropped=1 End If
If TargPos=45 Then RightTarget85.IsDropped=0 else RightTarget85.IsDropped=1 End If
If TargPos=46 Then RightTarget86.IsDropped=0 else RightTarget86.IsDropped=1 End If
If TargPos=47 Then RightTarget87.IsDropped=0 else RightTarget87.IsDropped=1 End If
If TargPos=48 Then RightTarget88.IsDropped=0 else RightTarget88.IsDropped=1 End If
If TargPos=49 Then RightTarget89.IsDropped=0 else RightTarget89.IsDropped=1 End If
If TargPos=50 Then RightTarget90.IsDropped=0 else RightTarget90.IsDropped=1 End If
If TargPos=51 Then RightTarget91.IsDropped=0 else RightTarget91.IsDropped=1 End If
If TargPos=52 Then RightTarget92.IsDropped=0 else RightTarget92.IsDropped=1 End If
If TargPos=53 Then RightTarget93.IsDropped=0 else RightTarget93.IsDropped=1 End If
If TargPos=54 Then RightTarget94.IsDropped=0 else RightTarget94.IsDropped=1 End If
If TargPos=55 Then RightTarget95.IsDropped=0 else RightTarget95.IsDropped=1 End If
If TargPos=56 Then RightTarget96.IsDropped=0 else RightTarget96.IsDropped=1 End If
If TargPos>56 and TargPos<59 Then RightTarget97.IsDropped=0 else RightTarget97.IsDropped=1 End If
If TargPos>58 and TargPos<62 Then RightTarget98.IsDropped=0 else RightTarget98.IsDropped=1 End If
If TargPos>61 and TargPos<65 Then RightTarget99.IsDropped=0 else RightTarget99.IsDropped=1 End If
If TargPos>64 Then RightTarget100.IsDropped=0 else RightTarget100.IsDropped=1 End If


If TargPos=70 Then TargDir=0 End If
If TargDir=1 Then TargPos=TargPos+1 End If
If TargDir=0 Then TargPos=TargPos-1 End If
End Sub


Sub Start_Game()
	If GameStart=1 and GameinProgress=1 then
    Randy=100
    lighta1.State=LightStateOn
	Score=0
	emreel1.ResetToZero()
	Advancer=1
EMReelCool.SetValue 0
EMReelWarm.SetValue 0
EMReelHot.SetValue 0
EMReelTorrid.SetValue 0
	MainTarget.IsDropped = False
	playsound SoundFXDOF("droptargetreset",112,DOFPulse,DOFContactors)
	If DesktopMode = True Then t1.visible=true
	t2.visible=false
	t3.visible=false
	t4.visible=false
	t5.visible=false
	t6.visible=false
	ht6.visible=false
	t7.visible=false
	t8.visible=false
	t9.visible=false
	t10.visible=false
	t11.visible=false
	ht11.visible=false
	t12.visible=false
	t13.visible=false
	t14.visible=false
	t15.visible=false
	t16.visible=false
	ht16.visible=false
	t17.visible=false
	t18.visible=false
	t19.visible=false
	t20.visible=false
	t21.visible=false
	ht21.visible=false
	t22.visible=false
	t23.visible=false
	t24.visible=false
	t25.visible=false
	top.visible=false
	htop.visible=false
	blowtimer.enabled=false
	light20.state=lightstateoff
	light7.state=lightstateon
	lightsl.state=lightstateoff
	lightsr.state=lightstateoff
	BGlight.state=lightstateoff
	BY1light.state=lightstateoff
	BY2light.state=lightstateoff
	BR1light.state=lightstateoff
	BR2light.state=lightstateoff
	lightkl.state=lightstateoff
	lightkr.state=lightstateon
	lightgl.state=1
	lightgr.state=1
	lighta.state=1
	lightb.state=1
	lighta2.state=lightstateoff
	lighta3.state=lightstateoff
	lighta4.state=lightstateoff
	lighta5.state=lightstateoff
	lightspecial.state=lightstateoff
	'MatchText.Text = "0"
	Ballstoplay=5
	BallstoPlay=BallsPerGame
	New_Ball
	End If
End Sub
' Quick Ball Sound V1.1 by STAT, stefanaustria
' -----------------------
Dim aX, aY, sX, RSound, SBall

Sub TriggerS_Timer()
	aX = int(SBall.VelX): aY = int(SBall.VelY)
	sX = -1.0: If int(Sball.X)>500 Then sX = 1.0
	If (aX>5 OR aY>5) AND Rsound = 0 Then
		RSound=int(RND*4)+1
		PlaySound "Roll"&RSound,0,1.0,sX,0.2 'replace sX to 0 for FullSound
	Elseif (aX<6 AND aY<6) AND Rsound > 0 Then
		StopSound "Roll"&Rsound
		Rsound = 0
	End If
End Sub

Sub TriggerS_Hit()
	DOF 122, DOFPulse
	Set SBall = Activeball
	TriggerS.TimerInterval = 100
	TriggerS.TimerEnabled = True
	PlaySound "roll1"
End Sub
' -----------------------

Sub BIP_hit
BIP=BIP+1
BallinLane=0
End Sub

Sub New_Ball
	DOF 118, DOFPulse
	'BallsinPlay=BallsinPlay+1
	'Kicker1.CreateBall.image="pinball"
MainTarget.IsDropped = False
LightA.state=1
LightB.state=1
	'Kicker1.kick 0,0
	If Shep=1 then
PlaySound "DrainShorter"
End If
If Shep=0 then
playsound "frigid"
Shep=1
End If
End Sub

Sub Drain1_hit
Drain1.DestroyBall
	PlaySoundAtVol "drain", drain, 1
End Sub
Sub Trigger1_hit()
Wall49.IsDropped=0
End Sub

Sub Drain_Hit()
	PlaySoundAtVol "drain", drain, 1
DOF 117, DOFPulse
TriggerS.TimerEnabled = False
BIP=BIP-1
TiltLight.SetValue 0
Tilt=0
	Drain.DestroyBall
If B2SOn then Controller.B2SSetTilt 0
Kicker2.CreateSizedBall(22)
Kicker2.Kick 90,2
	BallstoPlay=BallstoPlay-1
	If BallstoPlay=0 then
	EndofGame()
GameOver.SetValue 1
LeftFlipper.RotateToStart
RightFlipper.RotateToStart
motor.enabled=False
stopsound "motor"
targettimer.enabled=false
If B2SOn then Controller.B2SSetGameOver 1
	playsound "motorleer"
	End If
	If BallstoPlay>0 then
	Credittext.Text=Credits
	If B2SOn then Controller.B2SSetCredits Credits
	New_Ball
	End If
End Sub

Sub Triggerg1_hit()
If Tilt=0 and lightg1.state=1 then
Addscore (10)
call advance()
Playsound SoundFXDOF("1",141,DOFPulse,DOFChimes)
End If
End Sub

Sub Triggerg2_hit()
If Tilt=0 and lightg2.state=1 then
Addscore (10)
call advance()
Playsound SoundFXDOF("1",141,DOFPulse,DOFChimes)
End If
End Sub

Sub Triggerg3_hit()
If Tilt=0 and lightg3.state=1 then
Addscore (10)
call advance()
Playsound SoundFXDOF("1",141,DOFPulse,DOFChimes)
End If
End Sub

Sub Triggerg4_hit()
Addscore (10)
If Tilt=0 and lightg4.state=1 then
call advance()
Playsound SoundFXDOF("1",141,DOFPulse,DOFChimes)
End If
End Sub

Sub Triggerg5_hit()
Addscore (10)
If Tilt=0 and lightg5.state=1 then
call advance()
Playsound SoundFXDOF("1",141,DOFPulse,DOFChimes)
End If
End Sub

Sub Triggergl_hit()
If Tilt=0 and lightgl.state=1 then
playsound SoundFXDOF("10",142,DOFPulse,DOFChimes)
call advance()
Addscore (10)
scoretext.text=score
End If
End Sub

Sub Triggergr_hit()
If Tilt=0 and lightgr.state=1 then
playsound SoundFXDOF("10",142,DOFPulse,DOFChimes)
call advance()
Addscore (10)
scoretext.text=score
End If
End Sub


Sub BumperGreen_hit()
DOF 109, DOFPulse
If Tilt=0 then
If BGlight.State=0 then
Addscore (10)
scoretext.text=score
playsound SoundFXDOF("10",142,DOFPulse,DOFChimes)
End If
If BGlight.State=1 then
Addscore (100)
scoretext.text=score
playsound SoundFXDOF("100",143,DOFPulse,DOFChimes)
changopanjo()
If Omar=0 and lightgl.state=lightStateOff then
Omar=1
End If
If Omar=0 and Lightgl.state=LightStateOn then
End If
Matchnum=Matchnum+1
	If Matchnum=10 Then
	Matchnum=0
	End If

End If
Omar=0
End If
End Sub

Sub BumperY1_hit()
DOF 108, DOFPulse
If Tilt=0 then
changopanjo()
If Omar=0 and lightgl.state=lightStateOff then
Omar=1
End If


If BY1light.State=LightStateOff then
Addscore (1)
scoretext.text=score
playsound SoundFXDOF("1",141,DOFPulse,DOFChimes)
End If
If BY1light.State=LightStateOn then
Addscore (10)
scoretext.text=score
playsound SoundFXDOF("10",142,DOFPulse,DOFChimes)
End If
Omar=0
End If
End Sub

Sub BumperY2_hit()
DOF 111, DOFPulse
If Tilt=0 then
changopanjo()
If Omar=0 and lightgl.state=lightStateOff then
Omar=1
End If


If BY2light.State=LightStateOff then
Addscore (1)
scoretext.text=score
playsound SoundFXDOF("1",141,DOFPulse,DOFChimes)
End If
If BY2light.State=LightStateOn then
Addscore (10)
scoretext.text=score
playsound SoundFXDOF("10",142,DOFPulse,DOFChimes)
End If
Omar=0
End If
End Sub

Sub BumperR1_hit()
DOF 110, DOFPulse
If Tilt=0 then
changopanjo()
If Omar=0 and lightgl.state=lightStateOff then
Omar=1
End If

If BR1light.State=LightStateOff then
Addscore (1)
scoretext.text=score
playsound SoundFXDOF("1",141,DOFPulse,DOFChimes)
End If
If BR1light.State=LightStateOn then
Addscore (10)
scoretext.text=score
playsound SoundFXDOF("10",142,DOFPulse,DOFChimes)
End If
Omar=0
End If
End Sub

Sub BumperR2_hit()
DOF 107, DOFPulse
If Tilt=0 then
changopanjo()
If Omar=0 and lightgl.state=lightStateOff then
Omar=1
End If

If BR2light.State=LightStateOff then
Addscore (1)
scoretext.text=score
playsound SoundFXDOF("1",141,DOFPulse,DOFChimes)
End If
If BR2light.State=LightStateOn then
Addscore (10)
scoretext.text=score
playsound SoundFXDOF("10",142,DOFPulse,DOFChimes)
End If
Omar=0
End if
End Sub

Sub MainTarget_Hit()
playsoundAtVol SoundFXDOF("droptargetreset",112,DOFPulse,DOFContactors), ActiveBall, 1
If Tilt=0 then
changopanjo()
If Randy=100 then
Addscore (100)
scoretext.text=score
playsound SoundFXDOF("drop100",143,DOFPulse,DOFChimes)
MainTarget.IsDropped = True
LightA.State=LightStateOn
LightB.State=LightStateOn
'ChangerA_Hit()
'ChangerB_Hit()
End If
If Randy=200 then
Addscore (200)
scoretext.text=score
playsound "flourish"
MainTarget.IsDropped = True
LightA.State=LightStateOn
LightB.State=LightStateOn
'ChangerA_Hit()
'ChangerB_Hit()
End if
If Randy=300 then
Addscore (300)
scoretext.text=score
playsound "flourish"
MainTarget.IsDropped = True
LightA.State=LightStateOn
LightB.State=LightStateOn
'ChangerA_Hit()
'ChangerB_Hit()
End If
If Randy=400 then
Addscore (400)
scoretext.text=score
playsound "flourish"
MainTarget.IsDropped = True
LightA.State=LightStateOn
LightB.State=LightStateOn
'ChangerA_Hit()
'ChangerB_Hit()
End If
If Randy=500 then
Addscore (500)
scoretext.text=score
playsound "flourish"
MainTarget.IsDropped = True
LightA.State=LightStateOn
LightB.State=LightStateOn
End If
If Randy=700 then
Addscore (700)
scoretext.text=score
playsound "700"
If credits<9 then Credits=Credits+1:DOF 121, DOFOn:end if
Credittext.Text=Credits
If B2SOn then Controller.B2SSetCredits Credits
MainTarget.IsDropped = True
LightA.State=LightStateOn
LightB.State=LightStateOn
End If
End If
End Sub




Sub LeftTargets_Hit(idx)
    playsoundAtVol SoundFXDOF("1",141,DOFPulse,DOFChimes), ActiveBall, 1
	DOF 113, DOFPulse
If Tilt=0 then
Advance
    If LightA.State=LightStateOn and LightB.State=LightStateOff then
If MainTarget.IsDropped=True Then
PlaysoundAtVol SoundFXDOF("DropTargetReset",112,DOFPulse,DOFContactors), ActiveBall, 1
End If
    MainTarget.IsDropped = False
    LightB.State=LightStateOn
else
    LightA.State=LightStateOff
End If
    End If
    End Sub

Sub RightTargets_Hit(idx)
    playsoundAtVol SoundFXDOF("1",141,DOFPulse,DOFChimes), ActiveBall, 1
	DOF 114, DOFPulse
If Tilt=0 then
Advance
    If LightB.State=LightStateOn and LightA.State=LightStateOff then
If MainTarget.IsDropped=True Then
PlaysoundAtVol SoundFXDOF("DropTargetReset",112,DOFPulse,DOFContactors), ActiveBall, 1
End If
    MainTarget.IsDropped = False
    LightA.State=LightStateOn
else
    LightB.State=LightStateOff
End If
    End If
    End Sub


Sub Gate_hit
PlaysoundAtVol "gate", ActiveBall, 1
End Sub

Sub LeftOutlane_Hit()
DOF 115, DOFPulse
If Tilt=0 then
if Light7.state=Lightstateoff then
Addscore (100)
scoretext.text=score
playsound SoundFXDOF("100",143,DOFPulse,DOFChimes)
end if
If Light7.state=LightStateOn then
addscore (Randy)
if randy=100 then
playsound SoundFXDOF("100",143,DOFPulse,DOFChimes)
scoretext.text=score
end if
if randy>100 then
playsound "flourish"
scoretext.text=score
end if
End if
If Lightsl.state=lightStateOn then
If credits<9 then Credits=Credits+1:DOF 121, DOFOn:end if
Credittext.Text=Credits
If B2SOn then Controller.B2SSetCredits Credits
playsound SoundFXDOF("knock",119,DOFPulse,DOFKnocker)
DOF 120, DOFPulse
End If
End If
End Sub

Sub RightOutlane_Hit()
DOF 116, DOFPulse
If Tilt=0 then
if Light20.state=Lightstateoff then
Addscore (100)
scoretext.text=score
playsound SoundFXDOF("100",143,DOFPulse,DOFChimes)
end if
If Light20.state=LightStateOn then
addscore (Randy)
scoretext.text=score
if randy=100 then
playsound SoundFXDOF("100",143,DOFPulse,DOFChimes)
end if
if randy>100 then
playsound "flourish"
end if
end if
If Lightsr.state=lightStateOn then
If credits<9 then Credits=Credits+1:DOF 121, DOFOn:end if
Credittext.Text=Credits
If B2SOn then Controller.B2SSetCredits Credits
playsound SoundFXDOF("knock",119,DOFPulse,DOFKnocker)
DOF 120, DOFPulse
End If
End If
End Sub

Sub TempLevelLights_Timer
templevel=1
if ht6.visible=true then TempLevel=2
if ht11.visible=true then TempLevel=3
if ht16.visible=true then TempLevel=4
if ht21.visible=true then TempLevel=5
if htop.visible=true then TempLevel=6

If TempLevel=1 then EMReelFrigid.SetValue 1 End If
If TempLevel=2 then EMReelCool.SetValue 1 End If
If TempLevel=3 then EMReelWarm.SetValue 1 End If
If TempLevel=4 then EMReelHot.SetValue 1 End If
If TempLevel=5 then EMReelTorrid.SetValue 1 End If
If TempLevel=6 then EMreelblow.setvalue 1 End If
If TempLevel=3 then Lightgl.state=0:Lightgr.state=0: End If
If TempLevel=4 then Lightg1.state=0:Lightg5.state=0: End If
If TempLevel=5 then Lightg2.state=0:Lightg4.state=0: End If
If TempLevel=6 then Lightg3.state=0 End If


End Sub



Sub Advance
Advancer=Advancer+1
if Advancer>26 then Advancer=26
for i = 51 to 77
	If B2SOn then Controller.B2SSetData i,0
next
If B2SOn then Controller.B2SSetData Advancer+51,1
If Advancer=1 then
If DesktopMode = True Then t1.visible=true
End if
If Advancer=2 then
If DesktopMode = True Then t2.visible=true
End If
If Advancer=3 then
If DesktopMode = True Then t3.visible=true
End If
If Advancer=4 then
If DesktopMode = True Then t4.visible=true
End If
If Advancer=5 then
If DesktopMode = True Then t5.visible=true
End If
If Advancer=6 then
ht6.visible=true
If DesktopMode = True Then t6.visible=true
Lighta2.State=LightStateOn
If B2SOn then Controller.B2SSetData 82,1
Randy=200
BY2Light.State=LightStateOn
playsound "cool"
End If
If Advancer=7 then
If DesktopMode = True Then t7.visible=true
End If
If Advancer=8 then
If DesktopMode = True Then t8.visible=true
End If
If Advancer=9 then
If DesktopMode = True Then t9.visible=true
End If
If Advancer=10 then
If DesktopMode = True Then t10.visible=true
End If
If Advancer=11 then
ht11.visible=true
Lighta3.State=LightStateOn
Randy=300
If DesktopMode = True Then t11.visible=true
BR2light.State=LightStateOn
If B2SOn then Controller.B2SSetData 83,1
playsound "warm"
End If
If Advancer=12 then
If DesktopMode = True Then t12.visible=true
End If
If Advancer=13 then
If DesktopMode = True Then t13.visible=true
End If
If Advancer=14 then
If DesktopMode = True Then t14.visible=true
End If
If Advancer=15 then
If DesktopMode = True Then t15.visible=true
End If
If Advancer=16 then
ht16.visible=true
If DesktopMode = True Then t16.visible=true
lighta4.State=LightStateOn
Randy=400
BY1light.State=LightStateOn
stopsound "heatwave"
If B2SOn then Controller.B2SSetData 84,1
playsound "hot"
End If
If Advancer=17 then
If DesktopMode = True Then t17.visible=true
End If
If Advancer=18 then
If DesktopMode = True Then t18.visible=true
End If
If Advancer=19 then
If DesktopMode = True Then t19.visible=true
stopsound "heatwave"
playsound "heatwave"
End If
If Advancer=20 then
If DesktopMode = True Then t20.visible=true
End If
If Advancer=21 then
ht21.visible=true
If DesktopMode = True Then t21.visible=true
lighta5.State=LightStateOn
Randy=500
BR1light.State=LightStateOn
If B2SOn then Controller.B2SSetData 85,1
playsound "torrid"
End If
If Advancer=22 then
If DesktopMode = True Then t22.visible=true
End If
If Advancer=23 then
If DesktopMode = True Then t23.visible=true
End If
If Advancer=24 then
If DesktopMode = True Then t24.visible=true
End If
If Advancer=25 then
If DesktopMode = True Then t25.visible=true
End If
If Advancer=26 then
htop.visible=true
If DesktopMode = True Then top.visible=true
EMreelBlow.setvalue 1
light1.state=1
blowtimer.enabled=true
lightspecial.state=Lightstateon
Randy=700
if light7.state=lightstateon then
lightsl.state=Lightstateon
end if
if light20.state=lightstateon then
Lightsr.state=LightStateOn
end if
stopsound "heatwave"

playsound "blow"
End If
End Sub


Sub LeftSlingshot_Slingshot()
playsoundAtVol SoundFXDOF("slingshot",103,DOFPulse,DOFContactors), ActiveBall, 1
DOF 104, DOFPulse
If Tilt=0 then
If LightKl.State=LightStateOn then
Addscore (10)
scoretext.text=score
Playsound SoundFXDOF("10",142,DOFPulse,DOFChimes)
End If
IF Lightkl.State=LightStateOff then
Addscore (1)
playsound SoundFXDOF("1",141,DOFPulse,DOFChimes)
End If
End If
End Sub

Sub RightSlingshot_Slingshot()
playsoundAtVol SoundFXDOF("slingshot",105,DOFPulse,DOFContactors), ActiveBall, 1
DOF 106, DOFPulse
If Tilt=0 then
If LightKr.State=LightStateOn then
Addscore (10)
Playsound SoundFXDOF("10",142,DOFPulse,DOFChimes)
End If
IF Lightkr.State=LightStateOff then
Addscore (1)
playsound SoundFXDOF("1",141,DOFPulse,DOFChimes)
End If
End If
End Sub

Sub AddScore (points)
	Score=Score + points
	If B2SOn then Controller.B2SSetScorePlayer1 Score
	If score>1799 and Replay=0 then
		Playsound SoundFXDOF("knock",119,DOFPulse,DOFKnocker)
		DOF 120, DOFPulse
		If Credits<9 then Credits=Credits+1:DOF 121, DOFOn:end if
Credittext.text=credits
If B2SOn then Controller.B2SSetCredits Credits
		Replay=1
	End If
	If score>2699 and Replay=1 then
		Playsound SoundFXDOF("knock",119,DOFPulse,DOFKnocker)
		DOF 120, DOFPulse
		If Credits<9 then Credits=Credits+1:DOF 121, DOFOn:end if
Credittext.text=credits
If B2SOn then Controller.B2SSetCredits Credits
		Replay=2
	End If
	If score>3199 and Replay=2 then
		Playsound SoundFXDOF("knock",119,DOFPulse,DOFKnocker)
		DOF 120, DOFPulse
		If Credits<9 then Credits=Credits+1:DOF 121, DOFOn:end if
Credittext.text=credits
If B2SOn then Controller.B2SSetCredits Credits
		Replay=3
	End If
	If score>3899 and Replay=3 then
		Playsound SoundFXDOF("knock",119,DOFPulse,DOFKnocker)
		DOF 120, DOFPulse
		If Credits<9 then Credits=Credits+1:DOF 121, DOFOn:end if
Credittext.text=credits
If B2SOn then Controller.B2SSetCredits Credits
		Replay=4
	End If
emreel1.AddValue(points)
End Sub

sub changopanjo()
If Tilt=0 then
Desangel=Desangel+1
If Desangel=2 then
Desangel=0
End If
If Desangel=1 then
BGlight.state=lightstateon
light7.state=lightstateoff
light20.state=lightstateon
lightkl.state=lightstateon
lightkr.state=lightstateoff
if advancer>25 then
lightsl.state=lightstateoff
lightsr.state=lightstateon
end if
end if
If Desangel=0 then
BGlight.state=lightstateoff
light7.state=lightstateon
light20.state=lightstateoff
lightkl.state=lightstateoff
lightkr.state=lightstateon
if advancer>25 then
lightsl.state=lightstateon
lightsr.state=lightstateoff
end if
end if
End If
end sub


sub timer1_timer
	if timer1flag=0 then
		If B2SOn then Controller.B2SSetData 10,1
		timer1flag=1
	else
		If B2SOn then Controller.B2SSetData 10,0
		timer1flag=0
	end if
end sub

sub timer2_timer
	if timer2flag=0 then
		If B2SOn then Controller.B2SSetData 11,1
		timer2flag=1
	else
		If B2SOn then Controller.B2SSetData 11,0
		timer2flag=0
	end if
end sub
sub timer3_timer
	if timer3flag=0 then
		If B2SOn then Controller.B2SSetData 12,1
		timer3flag=1
	else
		If B2SOn then Controller.B2SSetData 12,0
		timer3flag=0
	end if
end sub
sub timer4_timer
	if timer4flag=0 then
		If B2SOn then Controller.B2SSetData 13,1
		timer4flag=1
	else
		If B2SOn then Controller.B2SSetData 13,0
		timer4flag=0
	end if
end sub
sub timer5_timer
	if timer5flag=0 then
		If B2SOn then Controller.B2SSetData 14,1
		timer5flag=1
	else
		If B2SOn then Controller.B2SSetData 14,0
		timer5flag=0
	end if
end sub
sub timer6_timer
	if timer6flag=0 then
		If B2SOn then Controller.B2SSetData 15,1
		timer6flag=1
	else
		If B2SOn then Controller.B2SSetData 15,0
		timer6flag=0
	end if
end sub

Sub Heatw_Exit()
	If B2SOn Then Controller.Stop
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

Function RndNum(min, max)
    RndNum = Int(Rnd() * (max-min + 1) ) + min ' Sets a random number between min and max
End Function

Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "Heatw" is the name of the table
  Dim tmp
  tmp = tableobj.y * 2 / Heatw.height-1
  If tmp > 0 Then
    AudioFade = Csng(tmp ^10)
  Else
    AudioFade = Csng(-((- tmp) ^10) )
  End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "Heatw" is the name of the table
  Dim tmp
  tmp = tableobj.x * 2 / Heatw.width-1
  If tmp > 0 Then
    AudioPan = Csng(tmp ^10)
  Else
    AudioPan = Csng(-((- tmp) ^10) )
  End If
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "Heatw" is the name of the table
    Dim tmp
    tmp = ball.x * 2 / Heatw.width-1
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

'*****************************************
'      JP's VP10 Rolling Sounds
'*****************************************

Const tnob = 6 ' total number of balls
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

Sub Rubbers_Hit(idx)
	Select Case Int(Rnd*3)+1
		Case 1 : PlaySound "rubber_hit_1", 0, Vol(ActiveBall)*400, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
		Case 2 : PlaySound "rubber_hit_2", 0, Vol(ActiveBall)*400, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
		Case 3 : PlaySound "rubber_hit_3", 0, Vol(ActiveBall)*400, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
	End Select
End Sub

Sub Slings_Hit(idx)
	Playsound "droptarg"
	Addscore (1)
	PlaySound (1)
End Sub

Sub LeftFlipper_Collide(parm)
 	RandomSoundFlipper()
End Sub

Sub RightFlipper_Collide(parm)
 	RandomSoundFlipper()
End Sub

Sub RandomSoundFlipper()
	Select Case Int(Rnd*3)+1
		Case 1 : PlaySound "flip_hit_1", 0, Vol(ActiveBall)*400, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
		Case 2 : PlaySound "flip_hit_2", 0, Vol(ActiveBall)*400, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
		Case 3 : PlaySound "flip_hit_3", 0, Vol(ActiveBall)*400, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
	End Select
End Sub





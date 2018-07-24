'
' SSSSSSSS PPPPPPPP AAAAAAAA CCCCCCCC EEEEEEEE    SSSSSSSS TTTTTTTT AAAAAAAA TTTTTTTT IIIIIIII OOOOOOOO NN    NN
' SS       PPPPPPPP AA    AA CC       EE          SS          TT    AA    AA    TT       II    OO    OO NNN   NN
' SS             PP AA    AA CC       EE          SS          TT    AA    AA    TT       II    OO    OO NNNN  NN
' SSSSSSSS PPPPPPPP AAAAAAAA CC       EEEEEEEE    SSSSSSSS    TT    AAAAAAAA    TT       II    OO    OO NN NN NN
'       SS PP       AA    AA CC       EE                SS    TT    AA    AA    TT       II    OO    OO NN  NNNN
'       SS PP       AA    AA CC       EE                SS    TT    AA    AA    TT       II    OO    OO NN   NNN
' SSSSSSSS PP       AA    AA CCCCCCCC EEEEEEEE    SSSSSSSS    TT    AA    AA    TT    IIIIIIII OOOOOOOO NN    NN
'
' Space Station / IPD No.2261 / December, 1987 / 4 Players
' http://www.ipdb.org/machine.cgi?id=2261
' VP9.16rev626 by MaX
' VP10 by nFozzy
' Version 1.11

'1.11
'Fixed ball reflections, updated a few scripts

'1.1 changelog
'different GI (and alternative incandescent GI)
'significant memory optimization
'new physics
'Ball shadow routine by Ninuzzu

' Thalamus 2018-07-24
' Table has its own "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' No special SSF tweaks yet.

Option explicit
Randomize

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

'**********************Options********************************

dim BallMass : BallMass = 1.65
dim SoundLevelMult, ReflectColor(2), FastFlips, DesktopMode : DesktopMode = Table1.ShowDT

SoundLevelMult = 1 'Increase table SFX (may cause some normalization)

GItype 2 'Initial GI type '0 = Random		'1 = #44 incandescent		'2 = White
'you can switch GI types in-game by pressing the right magnasave with the left magnasave held down

const SingleScreenFS = 0 '1 = VPX display 2 = Vpinmame display rotated

const BackGlassFlasher = 0 'Enable / Disable the backglass rocket flasher in DT mode (Default: 0)
Const HighGraphics = 1	'0 Enable / Disable wall reflections & ramp shadows (Default: 1)
Const BallReflections = 1 'Enable / Disable Ball Light Reflections (Default: 1)
const VPXdisplay = 1 'Enable/Disable VPX display. Disable for greater performance. (Default: 1)


'*************************************************************
'Debug stuff
Dim DebugFlippers : DebugFlippers = False
dim aDebugBoxes : aDebugBoxes = array(TbBounces, TB2, tbpl,tbWR, TbFlipper)
Sub DebugF(input)
	dim x
	if input = 5 then for each x in aDebugBoxes : x.visible = cBool(input) : next : exit sub
	input = cbool(input)
	if IsObject(fastflips) then FastFlips.Debug = input
'	TiltSol input
	destroyer.enabled = input
	Sw10.enabled = not input 'space station specific, disable trough
	if not input then Sw10.kick -10, 45 'space station specific
	for each x in aDebugBoxes : x.visible = input : next
	for each x in aDebugBoxes : x.TimerEnabled = input : next
end sub
debugf 0
'debugf 5 'boxes

sub GimmeF() : kl.createsizedballwithmass 25, ballmass : kl.kick 0, 5 : debugf 1 : end Sub
sub Gimmep() : kp.createsizedballwithmass 25, ballmass : kp.kick 0, 5 : debugf 1 : end Sub
sub Gimme() : kr.createsizedballwithmass 25, ballmass : kr.kick 0, 5 : debugf 1 : end Sub

'________________________________________________________
'    _____                      ______
'    /    '    /    ,             /
'---/__-------/------------__----/--------__----__---_/_-
'  /         /    /      /   )  /       /___)  (_ `  /
'_/_________/____/______/___/__/_______(___ __(__)__(_ __
'                      /
'Setup - 			  /
'Four Kickers named kiL, kL, kR, kiR placed in inlanes and on flippers
'Primitive FlipStick (make it a flat vertical stick with >1 opacity)
'Timer "FlipTest2"
'textbox "tb2"


dim FTSball : set FTSball = Nothing
dim FlipDir, FlipDelayV
Sub FTS(dir, input) 'hopefully more accurate flipper test sub
	debugf 1
'	FlipperLagCompensation 0	'remember to set this
	dim x : x = 0
	FlipDir = dir
	FlipDelayV = input
	Select Case dir
		case 0 : Set FTSball = kil.createsizedballwithmass(25, ballmass) : kil.kick 0, 0
		case 1 : Set FTSball = Kl.CreateSizedBallWithMass(25, ballmass) : Kl.Kick 2, 5 : SolBFlipper True : x = 2500
		case 2 : Set FTSball = kR.CreateSizedBallWithMass(25, ballmass) : Kr.Kick -2, 5 : SolBFlipper True : x = 2500
		case Else : Set FTSball = kir.createsizedballwithmass(25, ballmass) : kir.kick 0, 0 : flipdir = 3' : SolBFlipper True : x = 2000
	End Select
	if x > 0 then
		FlipDelayT1.Interval = x
		FlipDelayT1.Enabled = 1
	Else
		FlipTestV input	'fire flipper after a delay
	end if
End Sub

Sub FlipDelayT1_Timer(): SolBFlipper False : FlipTestV FlipDelayV : me.enabled = 0 : End Sub


Sub SolBFlipper(enabled)
	On Error Resume Next
	select case FlipDir
		case 2, 3 : FlipNF 1, enabled : x = (FTSball.x-RightFlipper.X) / (EndPointR - rightflipper.x)
		case else : FlipNF 0, enabled : x = (FTSball.x-LeftFlipper.X) / (EndPointL - LeftFlipper.x)
	end select
	if not enabled then exit sub
	FlipStick.x = FTSball.x : FlipStick.Visible = True
	x = RoundPercent(x)
	tb2.text = tb2.text & vbnewline & "%" & x' & vbnewline & "  flip: " & LeftFlipper.x & " ball:" & FTSball.x
End Sub

Function RoundPercent(input) : RoundPercent = mid(input, 1, 7)*100 : End Function	'round and mult by 100 for percentage

dim FlipInputV
Sub FlipTestV(input)
	tb2.text = "                          " & input & " MS"
	FlipTest2.Interval = input : FlipTestOn = True
	FlipTest2.Enabled = 1
End Sub

dim FlipTestOn
Sub FlipTest2_Timer()
	if FlipTestOn then
		SolBflipper True : me.interval = 100 : flipteston = False
	Else
		SolBflipper False
		me.enabled = 0
	end if
End Sub


'====ELASTICITY========================================================
' _______      ___       __       __        ______    _______  _______
'|   ____|    /   \     |  |     |  |      /  __  \  |   ____||   ____|
'|  |__      /  ^  \    |  |     |  |     |  |  |  | |  |__   |  |__
'|   __|    /  /_\  \   |  |     |  |     |  |  |  | |   __|  |   __|
'|  |      /  _____  \  |  `----.|  `----.|  `--'  | |  |     |  |
'|__|     /__/     \__\ |_______||_______| \______/  |__|     |__|
'======================================================================
'This script doesn't track ball angle like Jimmyfinger's script does so it may cause some weird ball movement on glancing shots
'Usage - Define a falloff line with two points, and then a string for debug purposes
'Falloffn x1,y1,x2,y2, DebugString
'(x = Input velocity, y = output Coef)
'Debug box "TBbounces"
'Usage: falloffsimple 1,1, 54,0.5, "Targets"
'TODO might add 3 point and 5 point envelopes

Sub FalloffSimple(X1, Y1, X2, Y2, DebugString)	'Two points
	Dim FinalSpeed : FinalSpeed = BallSpeed(ActiveBall) : if FinalSpeed < X1 then Exit Sub	'Cutoff Low
	Dim BounceCoef : BounceCoef = SlopeIt(FinalSpeed,X1,Y1,X2,Y2)
		if BounceCoef < Y2 then BounceCoef = Y2	: DebugString = DebugString & vbnewline & "Clamped" 'Clamp High
		activeball.velx = activeball.velx * BounceCoef
		activeball.vely = activeball.vely * BounceCoef

		DebugString = "FalloffSimple " & Debugstring & vbnewline
		FalloffDebugBox TBbounces, Finalspeed, BallSpeed(ActiveBall), BounceCoef, DebugString
End Sub

Sub FalloffDebugBox(object, input,output,Coef,debugstring)	'Debug Box
	'if not debugflippers then Exit Sub
	object.Text = Debugstring & round(input,4) & vbnewline & round(output,4) & vbnewline & "%" & round(coef,4)
	object.TimerEnabled = 1
End Sub
TBbounces.TimerInterval = 3000	'reset debug textbox after this interval
Sub TBbounces_Timer():me.timerenabled = 0 : me.text = Empty : End Sub
TBFlipper.TimerInterval = 5000	'reset debug textbox after this interval
Sub TBFlipper_Timer():me.timerenabled = 0 : me.text = Empty : End Sub


dim r1, r2, r3, r4 'posts
dim p1, p2, p3, p4 'pegs
dim s1, s2, s3, s4 'sleeves
dim t1, t2, t3, t4
'dim m1, m2, m3, m4	'flippers

p1 = 18 : p2 = 1 : p3 = 65 : p4 = 0.4
r1 = 18 : r2 = 1 : r3 = 60 : r4 = 0.25
s1 = 18 : s2 = 1 : s3 = 60 : s4 = 0.25
t1 = 2 : t2 = 1 : t3 = 50 :t4 =  0.4



'******************************************************************************
'     _______.  ______    __    __  .__   __.  _______      _______ ___   ___
'    /       | /  __  \  |  |  |  | |  \ |  | |       \    |   ____|\  \ /  /
'   |   (----`|  |  |  | |  |  |  | |   \|  | |  .--.  |   |  |__    \  V  /
'    \   \    |  |  |  | |  |  |  | |  . `  | |  |  |  |   |   __|    >   <
'.----)   |   |  `--'  | |  `--'  | |  |\   | |  '--'  |   |  |      /  .  \ 
'|_______/     \______/   \______/  |__| \__| |_______/    |__|     /__/ \__\ 
'
'******************************************************************************


'10.4 playsound args - name,loopcount,volume,pan,randompitch,pitch,UseExisting,Restart,Fade


'*********Sounds with falloffs*********
Sub Targets_Hit (idx)
	FalloffSimple t1, t2, t3, t4, "Targets"
	PlaySound SoundFX("target",DOFTargets), 0, LVL(0.2), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, Fade(ActiveBall)
	PlaySound SoundFX("targethit",0), 0, LVL(Vol(ActiveBall)*18 ), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, Fade(ActiveBall)
End Sub

Sub zCol_Sleeves_Hit() : FalloffSimple s1, s2, s3, s4, "Sleeves" : RandomSoundRubber : End sub

Sub Posts_Hit(idx)
	FalloffSimple r1, r2, r3, r4, "Posts"
	PlaySound RandomPost, 0, LVL(Vol(ActiveBall)*80 ), Pan(ActiveBall)*50, 0, Pitch(ActiveBall), 1, 0, Fade(ActiveBall)
End Sub

Sub Pegs_Hit(idx)
	FalloffSimple p1, p2, p3, p4, "Pegs"
	PlaySound RandomPost, 0, LVL(Vol(ActiveBall)*80 ), Pan(ActiveBall)*50, 0, Pitch(ActiveBall), 1, 0, Fade(ActiveBall)
End Sub



'*********Other Sounds*********
Sub RampEntry2_Hit()	'Left habitrail out of space station
	'Playsound "WireRamp", 0, LVL(0.3)
	WireRampOff 'Off of the plastic platform first
	WireRampOn False	'Quick hack arg :False = Habitrail, True = Plastic Ramp
End Sub

Sub RampSoundPlunge1_hit() : WireRampOn  False : End Sub
Sub RampSoundPlunge2_hit() : WireRampOff : WireRampOn True : End Sub	'Exit Habitrail, onto Mini PF
Sub RampSoundPlunge3_hit() : WireRampOn True : End Sub	'Exit Vuk, onto Mini PF

Sub RampEntry_Hit()
 	If activeball.vely < -8 then
		WireRampOn True
		'PlaySound "plasticrolling2", 0, LVL(1), Pan(ActiveBall), 0, Pitch(ActiveBall)*10, 1, 0
		PlaySound "ramp_hit2", 0, LVL(Vol(ActiveBall)), Pan(ActiveBall), 0, Pitch(ActiveBall)*10, 0, 0, Fade(activeball)
	Elseif activeball.vely > 3 then
		'StopSound "plasticrolling2"
		PlaySound "PlayfieldHit", 0, LVL(Vol(ActiveBall) ), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, Fade(activeball)
	End If
End Sub

'Sub RampSound0_Hit()	'mid plunge
'	If Activeball.velx > 0 then
'		Playsound "pmaReriW", 0, LVL(1)
'	Else
'		Playsound "WireRamp1", 0, LVL(1)
'	End If
'End Sub

Sub zApron_Hit (idx)
	PlaySound "woodhitaluminium", 0, LVL((Vol(ActiveBall))*10), Pan(ActiveBall)/4, 0, Pitch(ActiveBall), 1, 0,Fade(ActiveBall)
End Sub

Sub zPlastic_Hit (idx)
	PlaySound "Rubber_Hit_2", 0, LVL(Vol(ActiveBall)*30), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0,Fade(ActiveBall)
End Sub

Sub zInlanes_Hit (idx)
	PlaySound "MetalHit2", 0, LVL(Vol(ActiveBall)*5), Pan(ActiveBall)*55, 0, Pitch(ActiveBall), 1, 0,Fade(ActiveBall)
End Sub

Sub RampHit2_Hit()
	PlaySound "ramp_hit3", 0, LVL(Vol(ActiveBall) ), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0,Fade(ActiveBall)
End Sub

Sub Metals_Hit (idx)
	PlaySound "metalhit_medium", 0, LVL(Vol(ActiveBall)*30 ), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0,Fade(ActiveBall)
End Sub

Sub Gates_Hit (idx)
	PlaySound "gate", 0, LVL(Vol(ActiveBall) ), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0,Fade(ActiveBall)
End Sub

'Ramp drops using collision events
Sub col_UPF_Fall_hit():if FallSFX1.Enabled = 0 then playsound "drop_mono", 0, LVL(0.5), 0,0,0,0,0,Fade(ActiveBall) :FallSFX1.Enabled = 1	:end if :end sub'name,loopcount,volume,pan,randompitch
Sub FallSFX1_Timer():me.Enabled = 0:end sub

Sub Col_Rramp_Fall_Hit():if FallSFX2.Enabled = 0 then playsound "drop_mono", 0, LVL(0.5), 0.05,0,0,0,0,Fade(ActiveBall) :FallSFX2.Enabled = 1	:end if :end sub'name,loopcount,volume,pan,randompitch
Sub FallSFX2_Timer():me.Enabled = 0:end sub

Sub col_Lramp_Fall_Hit():if FallSFX3.Enabled = 0 then playsound "drop_mono", 0, LVL(0.5), -0.05,0,0,0,0,Fade(ActiveBall) :FallSFX3.Enabled = 1	:end if :end sub'name,loopcount,volume,pan,randompitch
Sub FallSFX3_Timer():me.Enabled = 0:end sub


Sub LeftSlingShot_Hit(): Playsound RandomBand, 0, LVL(Vol(ActiveBall)*100 ), Pan(ActiveBall)*50, 0, Pitch(ActiveBall), 1, 0,Fade(ActiveBall) : End Sub
Sub RightSlingShot_Hit(): Playsound RandomBand, 0, LVL(Vol(ActiveBall)*100 ), Pan(ActiveBall)*50, 0, Pitch(ActiveBall), 1, 0,Fade(ActiveBall) : End Sub

Sub Bands_Hit(idx) : Playsound RandomBand, 0, LVL(Vol(ActiveBall)*100 ), Pan(ActiveBall)*50, 0, Pitch(ActiveBall), 1, 0,Fade(ActiveBall) : End Sub
Sub RandomSoundRubber() : PlaySound RandomPost, 0, LVL(Vol(ActiveBall)*100 ), Pan(ActiveBall)*50, 0, Pitch(ActiveBall), 1, 0, Fade(ActiveBall) : End Sub



'SFX string functions -SFX - Posts 1-5, Bands 1-4 and 11,22,33,44
Function RandomPost() : RandomPost = "Post" & rndnum(1,5) : End Function

Function RandomBand()
		dim x : x = rndnum(1,4)
		if BallVel(activeball) > 30 then
			RandomBand = "Rubber" & x & x	'ex. Playsound "Band44"
		else
			RandomBand = "Rubber" & x	'ex. Playsound "Band4"
		End If
End Function

'Flipper collide sound
Sub LeftFlipper_Collide(parm) : RandomSoundFlipper() : End Sub
Sub RightFlipper_Collide(parm) : RandomSoundFlipper() : End Sub
Sub RandomSoundFlipper()
	dim x : x = RndNum(1,3)
	PlaySound "flip_hit_" & x, 0, LVL(Vol(ActiveBall)*50 ), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0,Fade(ActiveBall)
End Sub

' Ball Collision Sound
Sub OnBallBallCollision(ball1, ball2, velocity) : PlaySound("fx_collide"), 0, LVL(Csng(velocity) ^2 / 2000), Pan(ball1), 0, Pitch(ball1), 0, 0,Fade(ball1) : End Sub

'*****************************************
'      JP's VP10 Rolling Sounds
'*****************************************
Const tnob = 9 ' total number of balls
ReDim rolling(tnob)
InitRolling : Sub InitRolling : Dim i : For i = 0 to tnob : rolling(i) = False : Next : End Sub

Sub RollingTimer_Timer
    Dim BOT, b : BOT = GetBalls

	For b = UBound(BOT) + 1 to tnob	' stop the sound of deleted balls
		rolling(b) = False
		StopSound("tablerolling" & b)
	Next

	If UBound(BOT) = -1 Then Exit Sub	' exit the sub if no balls on the table

	For b = 0 to UBound(BOT)	' play the rolling sound for each ball
		If BallVel(BOT(b) ) > 1 AND BOT(b).z < 30 Then
			rolling(b) = True
			PlaySound("tablerolling" & b), -1, Vol(BOT(b) )^0.8, Pan(BOT(b) )*3, 0, BallPitch(BOT(b)), 1, 0,Fade(BOT(b))*3
		Else
			If rolling(b) = True Then
                StopSound("tablerolling" & b)
				rolling(b) = False
			End If
		End If
	Next
End Sub
Sub StopAllRolling() 	'call this at table pause!!!
	dim b : for b = 0 to tnob
		StopSound("tablerolling" & b)
		StopSound("RampLoop" & b)
		StopSound("wireloop" & b)
	next
end sub

'=====================================
'		Ramp Rolling SFX updates nf
'=====================================
'Ball tracking ramp SFX 1.0
'	Usage:
'- Setup hit events with WireRampOn True or WireRampOn False (True = Plastic ramp, False = Wire Ramp)
'- To stop tracking ball, use WireRampoff
'--	Otherwise, the ball will auto remove if it's below 30 vp units

'Example, from Space Station:
'Sub RampSoundPlunge1_hit() : WireRampOn  False : End Sub						'Enter metal habitrail
'Sub RampSoundPlunge2_hit() : WireRampOff : WireRampOn True : End Sub			'Exit Habitrail, enter onto Mini PF
'Sub RampEntry_Hit() : If activeball.vely < -10 then WireRampOn True : End Sub 	'Ramp enterance
dim RampMinLoops : RampMinLoops = 4
dim RampBalls(6,2)
'x,0 = ball x,1 = ID,	2 = Protection against ending early (minimum amount of updates)
'0,0 is boolean on/off, 0,1 unused for now
RampBalls(0,0) = False

dim RampType(6)	'Slapped together support for multiple ramp types... False = Wire Ramp, True = Plastic Ramp

Sub WireRampOn(input)  : Waddball ActiveBall, input : RampRollUpdate: End Sub
Sub WireRampOff() : WRemoveBall ActiveBall.ID	: End Sub

Sub Waddball(input, RampInput)	'Add ball
	dim x : for x = 1 to uBound(RampBalls)	'Check, don't add balls twice
		if RampBalls(x, 1) = input.id then
			if Not IsEmpty(RampBalls(x,1) ) then Exit Sub	'Frustating issue with BallId 0. Empty variable = 0
		End If
	Next

	For x = 1 to uBound(RampBalls)
		if IsEmpty(RampBalls(x, 1)) then
			Set RampBalls(x, 0) = input
			RampBalls(x, 1)	= input.ID
			RampType(x) = RampInput
			RampBalls(x, 2)	= 0
			'exit For
			RampBalls(0,0) = True
			RampRoll.Enabled = 1	 'Turn on timer
			'RampRoll.Interval = RampRoll.Interval 'reset timer
			exit Sub
		End If
		if x = uBound(RampBalls) then 	'debug
			Debug.print "WireRampOn error, ball queue is full: " & vbnewline & _
						 RampBalls(0, 0) & vbnewline & _
						 Typename(RampBalls(1, 0)) & " ID:" & RampBalls(1, 1) & "type:" & RampType(1) & vbnewline & _
						 Typename(RampBalls(2, 0)) & " ID:" & RampBalls(2, 1) & "type:" & RampType(2) & vbnewline & _
						 Typename(RampBalls(3, 0)) & " ID:" & RampBalls(3, 1) & "type:" & RampType(3) & vbnewline & _
						 Typename(RampBalls(4, 0)) & " ID:" & RampBalls(4, 1) & "type:" & RampType(4) & vbnewline & _
						 Typename(RampBalls(5, 0)) & " ID:" & RampBalls(5, 1) & "type:" & RampType(5) & vbnewline & _
						 " "
		End If
	next
End Sub

Sub WRemoveBall(ID)		'Remove ball
	dim ballcount : ballcount = 0
	dim x : for x = 1 to Ubound(RampBalls)
		if ID = RampBalls(x, 1) then 'remove ball
			Set RampBalls(x, 0) = Nothing
			RampBalls(x, 1) = Empty
			RampType(x) = Empty
			StopSound("RampLoop" & x)
			StopSound("wireloop" & x)
		end If
		'if RampBalls(x,1) = Not IsEmpty(Rampballs(x,1) then ballcount = ballcount + 1
		if not IsEmpty(Rampballs(x,1)) then ballcount = ballcount + 1
	next
	if BallCount = 0 then RampBalls(0,0) = False	'if no balls in queue, disable timer update
End Sub

Sub RampRoll_Timer():RampRollUpdate:End Sub

Sub RampRollUpdate()		'Timer update
	dim x : for x = 1 to uBound(RampBalls)
		if Not IsEmpty(RampBalls(x,1) ) then
			if BallVel(RampBalls(x,0) ) > 1 then ' if ball is moving, play rolling sound
				If RampType(x) then
					PlaySound("RampLoop" & x), -1, Vol(RampBalls(x,0) )*10, Pan(RampBalls(x,0) )*3, 0, BallPitchV(RampBalls(x,0) ), 1, 0,Fade(RampBalls(x,0) )'*3
					StopSound("wireloop" & x)
				Else
					StopSound("RampLoop" & x)
					PlaySound("wireloop" & x), -1, Vol(RampBalls(x,0) )*10, Pan(RampBalls(x,0) )*3, 0, BallPitch(RampBalls(x,0) ), 1, 0,Fade(RampBalls(x,0) )'*3
				End If
				RampBalls(x, 2)	= RampBalls(x, 2) + 1
			Else
				StopSound("RampLoop" & x)
				StopSound("wireloop" & x)
			end if
			if RampBalls(x,0).Z < 30 and RampBalls(x, 2) > RampMinLoops then	'if ball is on the PF, remove  it
				StopSound("RampLoop" & x)
				StopSound("wireloop" & x)
				Wremoveball RampBalls(x,1)
			End If
		Else
			StopSound("RampLoop" & x)
			StopSound("wireloop" & x)
		end if
	next
	if not RampBalls(0,0) then RampRoll.enabled = 0

End Sub


Sub tbWR_Timer()	'debug textbox
	me.text =	"on? " & RampBalls(0, 0) & " timer: " & RampRoll.Enabled & vbnewline & _
				 "1 " & Typename(RampBalls(1, 0)) & " ID:" & RampBalls(1, 1) & " type:" & RampType(1) & " Loops:" & RampBalls(1, 2) & vbnewline & _
				 "2 " & Typename(RampBalls(2, 0)) & " ID:" & RampBalls(2, 1) & " type:" & RampType(2) & " Loops:" & RampBalls(2, 2) & vbnewline & _
				 "3 " & Typename(RampBalls(3, 0)) & " ID:" & RampBalls(3, 1) & " type:" & RampType(3) & " Loops:" & RampBalls(3, 2) & vbnewline & _
				 "4 " & Typename(RampBalls(4, 0)) & " ID:" & RampBalls(4, 1) & " type:" & RampType(4) & " Loops:" & RampBalls(4, 2) & vbnewline & _
				 "5 " & Typename(RampBalls(5, 0)) & " ID:" & RampBalls(5, 1) & " type:" & RampType(5) & " Loops:" & RampBalls(5, 2) & vbnewline & _
				 "6 " & Typename(RampBalls(6, 0)) & " ID:" & RampBalls(6, 1) & " type:" & RampType(6) & " Loops:" & RampBalls(6, 2) & vbnewline & _
				 " "
End Sub

' *********************************************************************
'                      Ball & Sound Functions
' *********************************************************************
'10.4 playsound args - name,loopcount,volume,pan,randompitch,pitch,UseExisting,Restart,Fade

'**************** 3D Audio Vp10.4 Functions ****************
Function Fade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "table1" is the name of the table
	Dim tmp
    tmp = tableobj.y * 2 / table1.height-1
    If tmp > 0 Then
		Fade = Csng(tmp ^10)
    Else
        Fade = Csng(-((- tmp) ^10) )
    End If
End Function

Function FadeY(Y) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "table1" is the name of the table
	Dim tmp
    tmp = y * 2 / table1.height-1
    If tmp > 0 Then
		FadeY = Csng(tmp ^10)
    Else
        FadeY = Csng(-((- tmp) ^10) )
    End If
End Function

'**************** Other sound functions ****************
Function RndNum(min, max)
    RndNum = Int(Rnd() * (max-min + 1) ) + min ' Sets a random number between min and max
End Function

Function LVL(input) : LVL = Input * SoundLevelMult : End Function

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
    Vol = Csng(BallVel(ball) ^2 / 2000)
End Function

Function Vol2(ball1, ball2) ' Calculates the Volume of the sound based on the speed of two balls
    Vol2 = (Vol(ball1) + Vol(ball2) ) / 2
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "table1" is the name of the table
    Dim tmp : tmp = ball.x * 2 / Table1.width-1
    If tmp> 0 Then
        Pan = Csng(tmp ^10)
    Else
        Pan = Csng(-((- tmp) ^10) )
    End If
End Function

Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
    Pitch = BallVel(ball) * 20
End Function

Function BallVel(ball) 'Calculates the ball speed
    BallVel = INT(SQR(ball.VelX^2 + ball.VelY^2 + ball.VelZ^2) )
End Function

Function BallSpeed(ball) 'Calculates the ball speed
    BallSpeed = SQR(ball.VelX^2 + ball.VelY^2 + ball.VelZ^2)
End Function

'new
Function BallPitch(ball) ' Calculates the pitch of the sound based on the ball speed
    BallPitch = SlopeIt(BallVel(ball), 1, -1000, 60, 10000)
End Function
Function BallPitchV(ball) ' Calculates the pitch of the sound based on the ball speed Variation
    BallPitchV = SlopeIt(BallVel(ball), 1, -4000, 60, 7000)
End Function
'EOStimer (just switches elast falloff)
'============

'LFHM physics by wrd1972 and Rothbauer
dim EOSAngle,ElastFalloffUp,ElastFalloffDown

'This rules but it would be way better if it was elasticity and not elast falloff
ElastFalloffup = LeftFlipper.ElasticityFalloff
ElastFalloffdown = 0.7'0.7

'EOS angle
EOSAngle = 4

'Flipper EOS timer (HMLF)
dim LastAngle1, LastAngle2	'
Sub eostimer_Timer()	'use -1 timer interval for this?
	If LeftFlipper.CurrentAngle <> LastAngle1 then	'slight optimization
		If LeftFlipper.CurrentAngle < LeftFlipper.EndAngle + EOSAngle Then
			LeftFlipper.ElasticityFalloff = ElastFalloffup
		Else
			LeftFlipper.ElasticityFalloff = ElastFalloffdown	' This works for flippers :D
		End If
	End If
	If RightFlipper.CurrentAngle <> LastAngle2 then
		If RightFlipper.CurrentAngle > RightFlipper.EndAngle - EOSAngle Then
			RightFlipper.ElasticityFalloff = ElastFalloffup
		Else
			RightFlipper.ElasticityFalloff = ElastFalloffdown
		End If
	End If
	LastAngle1 = LeftFlipper.CurrentAngle
	LastAngle2 = RightFLipper.CurrentAngle
End Sub

'**************************************



Siderails.visible = Table1.ShowDT
F31.visible = Table1.ShowDT
F31.Visible = BackGlassFlasher

P_InstructionsFS.visible = Not Table1.ShowDT : P_InstructionsDT.visible = Table1.ShowDT

LoadVPM "01120100", "S11.VBS", 3.36
SetLocale(1033)

Const cGameName = "spstn_l5"

Const UseSolenoids = 1
Const UseLamps = 0
Const UseGI = 0
Const UseSync = 0
Const HandleMech = 0

' Standard Sounds
Const SCoin = "fx_Coin"

Dim bsTrough, bsLeftBallPopper, bsRightBallPopper, bsRightLock, bsLeftLock, dtdrop, dt3bank, mufo

Sub Table1_Init
	vpmInit Me
	With Controller
		.GameName = cGameName
		If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description:Exit Sub
		.SplashInfoLine = "Space Station (Williams 1987)"
'		.Games(cGameName).Settings.Value("rol") = 0
		.HandleKeyboard = 0
		.ShowTitle = 0
		.ShowDMDOnly = 1
		.ShowFrame = 0
		.HandleMechanics = 0
'		.Hidden = 0
		.SetDisplayPosition 0,0,GetPlayerHWnd 'if you can't see the DMD then uncomment this line
		On Error Resume Next
	'	Controller.SolMask(0) = 0
	'	vpmTimer.AddTimer 2000, "Controller.SolMask(0)=&Hffffffff'" 'ignore all solenoids - then add the timer to renable all the solenoids after 2 seconds
		Controller.Run GetPlayerHwnd
'        .DIP(0) = &H00		'Taxi
		If Err Then MsgBox Err.Description
		On Error Goto 0
	End With

	If SingleScreenFS = 2 and not Table1.ShowDT then
		Controller.Games(cGameName).Settings.Value("rol") = 1
	Else
		Controller.Games(cGameName).Settings.Value("rol") = 0
		Controller.Hidden = 1
	End If
	If VPXdisplay = 0 then
		Controller.Hidden = 0
	End If

  On Error Goto 0


	' Nudging
	vpmNudge.TiltSwitch = 9
	vpmNudge.Sensitivity = 0.25
    vpmNudge.TiltObj = Array(LeftSlingshot, RightSlingshot, bumper1, bumper2)',Bumper3)

	'Fastflips Pinmame callback bypass
	Set FastFlips = new cFastFlips
	with FastFlips
		.CallBackL = "SolLflipper"	'set flipper sub callbacks
		.CallBackR = "SolRflipper"	'...
		'.CallBackUL = "SolULflipper"'...(upper flippers, if needed)
		'.CallBackUR = "SolURflipper"'...
		.TiltObjects = True 'Optional, if True calls vpmnudge.solgameon automatically
		.Delay = 0			'Optional, if > 0 adds some compensation for solenoid jitter (occasional problem on Bram Stoker's Dracula)
		.Debug = False		'Debug, always-on flippers. Call FastFlips.Debug True or False in debugger to enable/disable.
	end with


	'sw41 Left Ball Popper (Top of Table)
	Set bsLeftBallPopper = New cvpmSaucer
	with bsLeftBallPopper
		.InitKicker LeftBallPopper, 41, 85, 85, 100 '82 'switch, direction, force, Zforce	'65
'		.InitExitVariance 2, 2
    End With

	'sw42 Right Ball Popper (Mega VUK)
	Set bsRightBallPopper = New cvpmSaucer
	with bsRightBallPopper
		.InitKicker RightBallPopper, 42, 85, 95, 100 'switch, direction, force, Zforce
		.InitExitVariance 0, 5
    End With

	'sw45 Right Lock
	Set bsRightLock = New cvpmSaucer
	with bsRightLock
		.InitKicker RightLock, 45, 0, 32, 0 'switch, direction, force, Zforce	'28
		.InitExitVariance 1, 2
    End With

	'sw47 Left Lock
	Set bsLeftLock = New cvpmSaucer
	with bsLeftLock
		'.InitKicker LeftLock, 47, 0, 32, 0 'switch, direction, force, Zforce
		.InitKicker LeftLock, 47, 0, 30, 0 'switch, direction, force, Zforce
		.InitExitVariance 5, 2
    End With

	'sw33 Drop Target 1-bank
	Set dtDrop = New cvpmDropTarget
	with dtDrop
		.InitDrop Array(sw33), Array(33)
		.InitSnd "fx_droptarget", "resetdrop"
'		.CreateEvents "dtDrop"
    End With

	'Drop Target 3-bank
	Set dt3bank = New cvpmDropTarget
	with dt3bank
		.InitDrop Array(sw57, sw58, sw59), Array(57, 58, 59)
		.InitSnd "", SoundFX("resetdrop",DOFDropTargets)
    End With

     ' Main Timer init
     PinMAMETimer.Interval = PinMAMEInterval
     PinMAMETimer.Enabled = 1
	 LampTimer.Enabled = 1

	' Init Kickback
    KickBack.Pullback

	'start trough
	sw13.CreateSizedBallWithMass 25, BallMass
	sw12.CreateSizedBallWithMass 25, BallMass
	sw11.CreateSizedBallWithMass 25, BallMass

	'Init animations
	LeftSling.Showframe 100
	LeftSlingArm.Showframe 100
	RightSling.Showframe 100
	RightSlingArm.Showframe 100
	BumperRing1.Showframe 100
	BumperRing2.Showframe 100
	BumperRing3.Showframe 100
	RubberAnim_1.Showframe 100
	RubberAnim_2.Showframe 100

	BallSearch	'init switches

End Sub
Sub table1_Paused:Controller.Pause = 1: StopAllRolling:End Sub
Sub table1_unPaused:Controller.Pause = 0:End Sub
Sub Table1_exit():Controller.Pause = False:Controller.Stop:End Sub
Sub Destroyer_Hit:Me.destroyball : End Sub

Sub zCol_Band_Anim1_Hit():RubberAnim_1.Playanim 0, (3*CGT/500):End Sub
Sub zCol_BandUPF_Hit():RubberAnim_2.Playanim 0, (3*CGT/500):End Sub

Sub Gate6_Hit(): activeball.velx = activeball.velx*0.95 : activeball.vely = activeball.vely*0.95 : End Sub	'switch and gate slow ball to left vuk
Sub Gate5_Hit(): activeball.velx = activeball.velx*0.95 : activeball.vely = activeball.vely*0.95 : End Sub

'New space station mech script script

SolCallback(16) = "Setlamp 0,"

Sub InitMotor():FlashSpeedUp(0) = 0.054 : End Sub
dim MotorPos : MotorPos = 90
Sub Motor(nr)	'Should go counter-clockwise...
	Select Case FadingLevel(nr)
		Case 4
			MotorSound False
			FadingLevel(nr) = 0
		Case 5
			MotorSound True
			MotorPos = MotorPos + FlashSpeedUp(nr) * CGT
			If MotorPos > 360 then MotorPos = MotorPos - 360
			If MotorPos < 0 then MotorPos = MotorPos + 360

			If MotorPos > 180 then Controller.Switch(52) = 1 Else Controller.Switch(52) = 0 End If
			If MotorPos > 90 and  MotorPos < 270 then Controller.Switch(53) = 1 Else Controller.Switch(53) = 0 End If
			UpdateUFO MotorPos
'			tb.text = MotorPos & vbnewline & " 1 " & Controller.Switch(52) & " 2 " & Controller.Switch(53)
	End Select
'	TBM.text = "Motorpos: " & MotorPos & vbnewline & "Speed: " & FlashSpeedUp(nr) & "Switch1: " & Controller.Switch(52) & VbNewline & "Switch2: " & Controller.Switch(53)
End Sub
dim MotorSoundON:MotorSoundON = 0

Sub MotorSound(aEnabled)
	If aEnabled Then
		If MotorSoundON then Exit Sub
		Playsound SoundFX("Motor",DOFgear), 0, LVL(1), Pan(SpaceStationToy), 0, 0, 1, 0, Fade(SpaceStationToy)
	Else
		Stopsound SoundFX("Motor",DOFgear)

	End If
	MotorSoundON = aEnabled
End Sub

' Update UFO
Sub UpdateUFO(aNewPos)
	SpaceStationToy.RotZ = (anewpos * -1)
	if AnewPos > 350 then GoLeft : Exit Sub
	If AnewPos < 10 Then GoLeft : Exit Sub
	if anewpos > 80 and anewpos < 100 then GoRight : Exit Sub
	if AnewPos > 170 and AnewPos < 190 then GoLeft : Exit Sub
	if AnewPos > 260 and AnewPos < 280 then GoRight : Exit Sub
	If ANewPos > 290 then GoRight : Exit Sub
	ToyClosed.isdropped = 0
End Sub
Sub GoLeft()
	toy1.isdropped=1:toy1a.isdropped=1:toy1b.isdropped=1:toy2.isdropped=0:toy2a.isdropped=0:toy2b.isdropped=0
'	tbm.text = "Go Left (0 180)"
	ToyClosed.isdropped = 1
End Sub

Sub GoRight()
	toy1.isdropped=0:toy1a.isdropped=0:toy1b.isdropped=0:toy2.isdropped=1:toy2a.isdropped=1:toy2b.isdropped=1
'	tbm.text = "Go Right (90 270)"
	ToyClosed.isdropped = 1
End Sub


dim CatchInput(1)
Sub Table1_KeyDown(ByVal keycode)
	If vpmKeyDown(keycode) Then Exit Sub
    If keycode = PlungerKey Then PlaySound SoundFx("plungerpull",0),0,LVL(1),0.25,0.25:Plunger.Pullback:End If
	If KeyCode = LeftFlipperKey then FastFlips.FlipL True' :  FastFlips.FlipUL True
	If KeyCode = RightFlipperKey then FastFlips.FlipR True' :  FastFlips.FlipUR True
	if Keycode = KeyRules then
		If DesktopMode Then
			P_InstructionsDT.PlayAnim 0, 3*CGT/500
		Else
			P_InstructionsFS.PlayAnim 0, 3*CGT/500
		End If
		Exit Sub
	End If
	if keycode = LeftMagnaSave then catchinput(0) = True
	if keycode = RightMagnaSave then catchinput(1) = True : if catchinput(0) and flashlevel(109) > 0 then gitype 5 : playsound "fx_relay_on", 0, LVL(0.1)
''	if keycode = 31 then Testtu 'test backhand shot
'
'	if keycode = 203 then TestL 'test backhand shot	'leftarrow
'	if keycode = 200 then TestM 'test backhand shot	'Uparrow
'	if keycode = 205 then TestR 'test backhand shot	'Right Arrow

'	If keycode = LeftFlipperKey Then :FlippersEnabled = True: flipnf 0, 1: exit sub:End If	'debug always-on flippers
'	If keycode = RightFlipperKey Then :FlippersEnabled = True: flipnf 1, 1: exit sub:End If	'debug always-on flippers
End Sub

Sub Table1_KeyUp(ByVal keycode)
    If vpmKeyUp(keycode) Then Exit Sub
    If keycode = PlungerKey Then
		Plunger.Fire
		if BallInPlunger then
			PlaySound SoundFX("plunger3",0),0,LVL(1),0.05,0.02,0,0,0,FadeY(1900)
		Else
			PlaySound SoundFX("plunger",0),0,LVL(0.8),0.05,0.02,0,0,0,FadeY(1900)
		end if
	End If
	if Keycode = KeyRules then
		If DesktopMode Then
			P_InstructionsDT.ShowFrame 0
		Else
			P_InstructionsFs.ShowFrame 0
		End If
		Exit Sub
	End If
	if keycode = LeftMagnaSave then catchinput(0) = False
	if keycode = RightMagnaSave then catchinput(1) = False

	If KeyCode = LeftFlipperKey then FastFlips.FlipL False' :  FastFlips.FlipUL False
	If KeyCode = RightFlipperKey then FastFlips.FlipR False' :  FastFlips.FlipUR False
End Sub

'VP10.2 Style OBJ Animations note

'VP10 animation OBJ squences only play back properly at 60FPS.
'So playback speed must be compensated for in script to play back at the proper speed regardless of framerate.
'x = CGT y = Playback Speed
'y = 3x/50


Sub TestAnims
	dim x
	x = 3*CGT/500
	LeftSling.playanim 0, x
	LeftSlingArm.playanim 0, x
	RightSling.playanim 0, x
	RightSlingArm.playanim 0, x
	BumperRing1.playanim 0, x
	BumperRing2.playanim 0, x
	BumperRing3.playanim 0, x
	RubberAnim_1.Playanim 0, x
	RubberAnim_2.Playanim 0, x
'	lefta.playanim 0, (3*CGT/500)
End Sub

'Trough Handling
'==============
SolCallback(1)	   = "TroughIn"       ' Drain
SolCallback(2)     = "TroughOut"      ' Ball Release
Sub TroughIn(enabled)
	if Enabled then
		sw10.Kick 60, 16 :
		Tgate.Open = True
		If sw10.BallCntOver > 0 Then Playsound SoundFX("Trough1",DOFcontactors), 0, LVL(0.5), 0.01 : End If
	Else
		Tgate.Open = False
		BallSearch
	End If
end Sub
Sub TroughOut(enabled): if Enabled then sw11.Kick 58, 8 :Playsound SoundFX("BallRelease",DOFcontactors), 0, LVL(0.4), 0.02:End If: end Sub

'Trough Switches
'Sub sw10_hit():controller.Switch(10) = 1 : Playsound "Trough2", 0, 0.2, 0: End Sub	'Drain
Sub sw10_hit():controller.Switch(10) = 1 : End Sub	'Drain
Sub TroughSFX_Hit(): Playsound "Trough2", 0, LVL(0.2), 0:: End Sub
Sub Sw13_hit():controller.Switch(13) = 1 : UpdateTrough : End Sub
Sub Sw12_hit():controller.Switch(12) = 1 : UpdateTrough : End Sub
Sub Sw11_hit():controller.Switch(11) = 1 : UpdateTrough : End Sub

Sub sw10_UnHit():controller.Switch(10) = 0 : UpdateTrough : End Sub
Sub Sw13_UnHit():controller.Switch(13) = 0 : UpdateTrough : End Sub
Sub Sw12_UnHit():controller.Switch(12) = 0 : UpdateTrough : End Sub
Sub Sw11_UnHit():controller.Switch(11) = 0 : UpdateTrough : End Sub

Sub UpdateTrough: TroughTimer.enabled = 1 : TroughTimer.Interval = 200: end sub

Sub TroughTimer_Timer()
	If sw11.BallCntOver = 0 then sw12.kick 58, 12
	If sw12.BallCntOver = 0 then sw13.kick 58, 12
End Sub

Sub BallSearch()	'In case of hard pinmame reset. Called by PF solenoids firing empty.
	if Sw10.BallCntOver > 0 then controller.Switch(10) = 1 else controller.Switch(10) = 0
	if Sw13.BallCntOver > 0 then controller.Switch(13) = 1 else controller.Switch(13) = 0
	if Sw12.BallCntOver > 0 then controller.Switch(12) = 1 else controller.Switch(12) = 0
	if Sw11.BallCntOver > 0 then controller.Switch(11) = 1 else controller.Switch(11) = 0
	If MotorPos > 180 then Controller.Switch(52) = 1 Else Controller.Switch(52) = 0 End If
	If MotorPos > 90 and  MotorPos < 270 then Controller.Switch(53) = 1 Else Controller.Switch(53) = 0 End If
	UpdateUFO MotorPos
End Sub
'==============SWITCHES==============
'====================================
'====================================

Sub sw33_Dropped : dtDrop.hit 1 : End Sub

'Sweep Bank Handling
Sub	sw57w_Hit : dt3bank.hit 1 : Sw57p.IsDropped = True : Sw57.IsDropped = True : Me.Enabled = 0 : Playsound SoundFX("DropTarget",DOFDropTargets), 0, LVL(1), 0.005: End Sub
Sub	sw58w_Hit : dt3bank.hit 2 : Sw58p.IsDropped = True : Sw58.IsDropped = True : Me.Enabled = 0 : Playsound SoundFX("DropTarget",DOFDropTargets), 0, LVL(1), 0.005: End Sub
Sub	sw59w_Hit : dt3bank.hit 3 : Sw59p.IsDropped = True : Sw59.IsDropped = True : Me.Enabled = 0 : Playsound SoundFX("DropTarget",DOFDropTargets), 0, LVL(1), 0.005: End Sub

'SolCallback(6)	   = "dt3bank.SolDropUp"       ' 3 Drop Target Bank Reset
SolCallback(6)	   = "ResetDTbank"       ' 3 Drop Target Bank Reset
Sub ResetDTbank(Enabled)
	dt3bank.SolDropUp Enabled
		if Enabled then
'			tb.text = "DT RESET"
			sw57p.Isdropped = False
			sw58p.IsDropped = False
			sw59p.IsDropped = False
			sw57.Isdropped = False
			sw58.IsDropped = False
			sw59.IsDropped = False
			sw57w.Enabled = True
			sw58w.Enabled = True
			sw59w.Enabled = True
		End If
End Sub

'Upper Rollovers
Sub sw14_Hit:Controller.Switch(14) = 1:End Sub
Sub sw14_UnHit:Controller.Switch(14) = 0:End Sub
Sub sw15_Hit:Controller.Switch(15) = 1:End Sub
Sub sw15_UnHit:Controller.Switch(15) = 0:End Sub
Sub sw16_Hit:Controller.Switch(16) = 1:End Sub
Sub sw16_UnHit:Controller.Switch(16) = 0:End Sub

'Outlane Switches
Sub sw17_Hit:Controller.Switch(17) = 1:End Sub
Sub sw17_UnHit:Controller.Switch(17) = 0:End Sub
Sub sw32_Hit:Controller.Switch(32) = 1:End Sub
Sub sw32_UnHit:Controller.Switch(32) = 0:End Sub

'Targets
Sub sw18_Hit:vpmTimer.PulseSw 18:End Sub
Sub sw19_Hit:vpmTimer.PulseSw 19:End Sub
Sub sw20_Hit:vpmTimer.PulseSw 20:End Sub
Sub sw21_Hit:vpmTimer.PulseSw 21:End Sub
Sub sw22_Hit:vpmTimer.PulseSw 22:End Sub
Sub sw23_Hit:vpmTimer.PulseSw 23:End Sub
Sub sw24_Hit:vpmTimer.PulseSw 24:End Sub
Sub sw25_Hit:vpmTimer.PulseSw 25:End Sub
Sub sw26_Hit:vpmTimer.PulseSw 26:End Sub
Sub sw27_Hit:vpmTimer.PulseSw 27:End Sub
Sub sw28_Hit:vpmTimer.PulseSw 28:End Sub
Sub sw29_Hit:vpmTimer.PulseSw 29:End Sub
Sub sw30_Hit:vpmTimer.PulseSw 30:End Sub
Sub sw31_Hit:vpmTimer.PulseSw 31:End Sub

'Big target
Sub sw35_Hit:vpmTimer.PulseSw 35:PlaySound "target", 0, LVL(0.1), 0.3:End Sub

'Rollunder
Sub sw37_Hit:Controller.Switch(37) = 1:End Sub
Sub sw37_UnHit:Controller.Switch(37) = 0:End Sub

'USA rollovers
Sub sw38_Hit:Controller.Switch(38) = 1:End Sub
Sub sw38_UnHit:Controller.Switch(38) = 0:End Sub
Sub sw39_Hit:Controller.Switch(39) = 1:End Sub
Sub sw39_UnHit:Controller.Switch(39) = 0:End Sub
Sub sw40_Hit:Controller.Switch(40) = 1:End Sub
Sub sw40_UnHit:Controller.Switch(40) = 0:End Sub

'Shooter Lane
Sub sw43_Hit:Controller.Switch(43) = 1:End Sub
Sub sw43_UnHit:Controller.Switch(43) = 0:End Sub

dim BallInPlunger :BallInPlunger = False
sub PlungerLane_hit():ballinplunger = True: End Sub
Sub PlungerLane_unhit():BallInPlunger = False: End Sub

'Right Lock Entry
Sub sw46_Hit:Controller.Switch(46) = 1:End Sub
Sub sw46_UnHit:Controller.Switch(46) = 0:End Sub

'Left Lock
Sub sw47_Hit:Controller.Switch(47) = 1:End Sub
Sub sw47_UnHit:Controller.Switch(47) = 0:End Sub

'Left Lock Entry (On the Ramp)
Sub sw48_Hit:Controller.Switch(48) = 1:End Sub
Sub sw48_UnHit:Controller.Switch(48) = 0:End Sub

'sw50 10 point (Top right)
Sub zCol_Band_Sw50_Hit():vpmTimer.PulseSw 50:End Sub

'sw54 10 point (Lwr right)
Sub zCol_Band_Sw54_Hit():vpmTimer.Pulsesw 54:End Sub

' Bumpers
Sub Bumper1_Hit:vpmTimer.PulseSw 60:PlaySound SoundFX("LeftBumper_Hit",DOFContactors), 	0, LVL(1), -0.01, 0.25:	BumperRing1.playanim 0, (3*CGT/500) * 0.97:End Sub
Sub Bumper2_Hit:vpmTimer.PulseSw 61:PlaySound SoundFX("RightBumper_Hit",DOFContactors), 0, LVL(1), 0.01, 0.25:	BumperRing2.playanim 0, (3*CGT/500) * 0.97:End Sub
Sub Bumper3_Hit:vpmTimer.PulseSw 62:PlaySound SoundFX("TopBumper_Hit",DOFContactors), 	0, LVL(1), 0, 0.25:		BumperRing3.playanim 0, (3*CGT/500) * 0.97:End Sub

'Drop-Targets
Sub sw33_Hit():dtDrop.Hit 1: Playsound SoundFX("DropTarget",DOFDropTargets), 0, LVL(1), 0.02,0,0,0,0,FadeY(1900):End Sub

'Sling Shots
Sub LeftSlingShot_Slingshot

	vpmTimer.PulseSw 63
    PlaySound SoundFX("LeftSlingShot",DOFContactors),0,LVL(1),-0.01,0.05
	LeftSling.playanim 0, (3*CGT/500) * 1.25
	LeftSlingArm.playanim 0, (3*CGT/500) * 1.25
End Sub

Sub RightSlingShot_Slingshot
	vpmTimer.PulseSw 64
    PlaySound SoundFX("RightSlingShot",DOFContactors), 0, LVL(1), 0.01, 0.05
	RightSling.playanim 0, (3*CGT/500) * 1.25
	RightSlingArm.playanim 0, (3*CGT/500) * 1.25
End Sub

'*********
'Solenoids

SolCallback(8)	   = "dtdrop.SolDropUp"		' Single Drop Target Reset
SolCallback(13)    = "SolKickBack"    ' Left Re-Entry Kickback

SolCallback(3)     = "LeftVuk"	' Left  Ball Popper
SolCallback(4)	   = "RightVuk"	' Right Ball Popper

Sub LeftVuk(enabled)
	If Enabled then
		If Toy2.isdropped then
			bsLeftBallPopper.InitKicker LeftBallPopper, 41, 85, 65, 100 '82 'switch, direction, force, Zforce	'65
		Else
			bsLeftBallPopper.InitKicker LeftBallPopper, 41, 85, 85, 100 '82 'switch, direction, force, Zforce	'65
		End If
		bsLeftBallPopper.ExitSol_On
		if LeftBallPopper.BallCntOver > 0 Then
			Playsound SoundFX("fx_VukOut2",DOFContactors), 1, LVL(5), 0
		Else
			Playsound SoundFX("DiverterLeft",DOFContactors), 1, LVL(0.25), 0
			BallSearch
		End If
	End If
End Sub

Sub RightVuk(enabled)
	If Enabled then
		bsRightBallPopper.ExitSol_On
		if RightBallPopper.BallCntOver > 0 Then
			Playsound SoundFX("fx_VukOut",DOFContactors), 1, LVL(5), 0.01
		Else
			Playsound SoundFX("DiverterRight",DOFContactors), 1, LVL(0.25), 0.1
			BallSearch
		End If
	End If
End Sub

SolCallback(17)    = "RightLockOut"   ' Right lock kickback
SolCallback(32)    = "LeftLockOut"	' Left lock Kickback

Sub LeftLockOut(enabled)
	If Enabled then
		bsLeftLock.ExitSol_On
		SafetyL.Enabled = 0: SafetyL.Kick 0, 22 	'bugfix - prevent 2 balls from getting stuck in this lock
		if LeftLock.BallCntOver > 0 Then
			Playsound SoundFX("Kicker_Release",DOFContactors), 1, LVL(1), -0.02
		Else
			Playsound SoundFX("DiverterLeft",DOFContactors), 1, LVL(0.25), -0.02
			BallSearch
		End If
	End If
End Sub

Sub RightLockOut(enabled)
	If Enabled then
		bsRightLock.ExitSol_On
		SafetyR.Enabled = 0: SafetyR.Kick 0, 22	'bugfix - prevent 2 balls from getting stuck in this lock
		if RightLock.BallCntOver > 0 Then
			Playsound SoundFX("Kicker_Release",DOFContactors), 1, LVL(1), 0.01
		Else
			Playsound SoundFX("DiverterLeft",DOFContactors), 1, LVL(0.25), 0.01
			BallSearch
		End If
	End If
End Sub


'PlaySound "name",loopcount,volume,pan,randompitch
Sub LeftBallPopper_Hit():bsLeftBallPopper.Addball me : Playsound SoundFX("Kicker_Hit",DOFContactors), 1, LVL(0.5), 0,0,0,0,0,FadeY(70): End Sub
Sub RightBallPopper_Hit():bsRightBallPopper.AddBall me : Playsound SoundFX("Kicker_Hit",DOFContactors), 1, LVL(0.5), 0.005,0,0,0,0,FadeY(70) : End Sub

Sub LeftLock_Hit():bsLeftLock.AddBall me : Playsound SoundFX("TroughLock",DOFContactors), 1, LVL(0.8), -0.02,0,0,0,0,Fade(me) : SafetyL.Enabled = True : End Sub
Sub RightLock_Hit():bsRightLock.AddBall me : Playsound SoundFX("TroughLock",DOFContactors), 1, LVL(0.8), 0.01,0,0,0,0,Fade(me) : SafetyR.Enabled = True : End Sub

SolCallback(7) = "KnockerSol"
Sub KnockerSol(enabled) : If Enabled then Playsound SoundFX("Knocker",DOFKnocker), 0, LVL(0.5) : End If : End Sub

Sub SolKickBack(enabled)
    If enabled Then
       Kickback.Fire
       PlaySound SoundFX("DiverterLeft",DOFContactors), 0, LVL(1), -0.02
    Else
       KickBack.PullBack
    End If
End Sub



'*********
' FastFlips NF 'Pre-solid state flippers lag reduction
'*********
SolCallback(23)		= "FastFlips.Flippers_On"


'********************
' Special JP Flippers 'Legacy Flippers using callbacks
'********************
'SolCallback(sLRFlipper) = "SolRFlipper"
'SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
	If Enabled Then
		playsound SoundFX("FlipperUpLeft",DOFFlippers), 0, LVL(1), -0.01	'flip
		LeftFlipper.RotateToEnd
		ProcessballsL
	Else
		LeftFlipper.RotateToStart
		if LeftFlipper.CurrentAngle = LeftFlipper.StartAngle then
			playsound SoundFX("FlipperDown",DOFFlippers), 0, 0
		else
			playsound SoundFX("FlipperDown",DOFFlippers), 0, LVL(1), 0.01	'return
		end if
	End If
End Sub

Sub SolRFlipper(Enabled)
	If Enabled Then
		playsound SoundFX("FlipperUpLeft",DOFFlippers), 0, LVL(1), 0.01	'flip
		ProcessballsR
		RightFlipper.RotateToEnd
	Else
		RightFlipper.RotateToStart
		if RightFlipper.CurrentAngle = RightFlipper.StartAngle then
			playsound SoundFX("FlipperDown",DOFFlippers), 0, 0
		else
			playsound SoundFX("FlipperDown",DOFFlippers), 0, LVL(1), 0.01	'return
		end if

	End If
End Sub


'================VP10 Fading Lamps Script

Dim LampState(200), FadingLevel(200)
Dim FlashSpeedUp(200), FlashSpeedDown(200), FlashMin(200), FlashMax(200), FlashLevel(200)

dim CGT	: dim InitFadeTime(0) : InitFadeTime(0) = 0	'compensated game time

Sub LampTimer_Timer()  'set up for -1 timers but will work with whatever
	cgt = gametime - InitFadeTime(0)
    Dim chgLamp, num, chg, ii
    chgLamp = Controller.ChangedLamps
    If Not IsEmpty(chgLamp) Then
        For ii = 0 To UBound(chgLamp)
            LampState(chgLamp(ii, 0) ) = chgLamp(ii, 1)       'keep the real state in an array
            FadingLevel(chgLamp(ii, 0) ) = chgLamp(ii, 1) + 4 'actual fading step
        Next

    End If

    UpdateLamps
	InitFadeTime(0) = gametime
End Sub


Sub InitLamps()
    Dim x
    For x = 0 to 200
        LampState(x) = 0        ' current light state, independent of the fading level. 0 is off and 1 is on
        FadingLevel(x) = 4      ' used to track the fading state
        FlashSpeedUp(x) = 0.015   ' faster speed when turning on the flasher
        FlashSpeedDown(x) = 0.009 ' slower speed when turning off the flasher
        FlashMax(x) = 1         ' the maximum value when on, usually 1
        FlashMin(x) = 0         ' the minimum value when off, usually 0
		'Starting this at >0 is important to prevent initial stuttering
        FlashLevel(x) = 0.01       ' the intensity of the flashers, usually from 0 to 1
    Next
	for x = 109 to 112	'GI
        FlashSpeedUp(x) = 0.01125
        FlashSpeedDown(x) = 0.01125'0.0079
	next
	for x = 115 to 200	'flashers
        FlashSpeedUp(x) = 0.01875 '0.4  ' faster speed when turning on the flasher
        FlashSpeedDown(x) = 0.01	'0.2 ' slower speed when turning off the flasher
'        FlashSpeedDown(x) = 0.01125	'0.2 ' slower speed when turning off the flasher
	next
	initmotor()
	FlashLevel(109) = 1	'start with GI on
	FlashLevel(111) = 1	'start with GI on
	Setlamp 109, 1	'GI
	SetLamp 110, 0	'GI Green
	SetLamp 111, 1	'Gi Bumpers
	for each x in ReflectionsGI
		x.state = BallReflections
		x.visible = BallReflections
	Next
	'For each x in ReflectionsGIG
	'	x.state = BallReflections
	'	x.visible = BallReflections
	'Next
	For each x in ReflectionsGIB
		x.state = BallReflections
		x.visible = BallReflections
	next

	gi_left.visible = highgraphics
	gig_left.visible = highgraphics
	gi_back.visible = highgraphics
	gig_back.visible = highgraphics
	gi_right.visible = highgraphics
	gig_right.visible = highgraphics
'	f15w.visible = highgraphics
	f26w.visible = highgraphics
	f27w.visible = highgraphics
	f28w.visible = highgraphics
	f29w.visible = highgraphics
	shadowtest.visible = highgraphics
	shadowtestp.visible = highgraphics
	fbg1.visible = highgraphics
	fbg2.visible = highgraphics
	fbg3.visible = highgraphics
	fbg4.visible = highgraphics
	fbg5.visible = highgraphics
	fbg6.visible = highgraphics

	UpdateLamps
End Sub

Sub Tweakinserts(input1, input2)
	dim x, x2
	select case Input1
		case 1
			for each x in ColInserts
				x.Opacity = input2
			next
			x2 = "Opacity"
		case 2
			for each x in ColInserts
				x.ModulatevsAdd = input2
			next
			x2 = "Mod"
		Case 3
			For x = 0 to 200
				FlashSpeedUp(x) = input2   ' faster speed when turning on the flasher
			next

			x2 = "Fadespeedup"
		Case 4
			For x = 0 to 200
				FlashSpeedDown(x) = input2 ' slower speed when turning off the flasher
			next

			x2 = "FadespeedDown"
	End Select

	tb.text = "tweaked " & x2 & vbnewline & "Opacity:" & L34.Opacity & vbnewline & "Mod:" & l34.ModulatevsAdd & _
		vbnewline & "fadeup:" & FLashSpeedUp(34) & vbnewline & "fadeDown:" & FLashSpeedDown(34)

End Sub
Sub Tweakref(input1, input2)
	dim x, x2
	select case Input1
		case 1
			for each x in ReflectionsGI
				x.Intensity = Input2
			next
			x2 = "Intensity"
		case 2
			for each x in ReflectionsGI
				x.BulbModulatevsAdd = input2
			next
			x2 = "Mod"
		Case 3
			for each x in ReflectionsGI
				x.Falloff = input2
			next
			x2 = "Falloff"
		Case 4
			for each x in ReflectionsGI
				x.Falloffpower = input2
			next
			x2 = "FalloffPower"
		Case 5
			for each x in ReflectionsGI
				x.bulbhaloheight = input2
			next
			x2 = "height"
	End Select

	tb.text = "tweaked " & x2' & vbnewline & "Opacity:" & L34.Opacity & vbnewline & "Mod:" & l34.ModulatevsAdd & _
'		vbnewline & "fadeup:" & FLashSpeedUp(34) & vbnewline & "fadeDown:" & FLashSpeedDown(34)

End Sub

'	for each x in ReflectionsGI
'		x.state = BallReflections
'		x.visible = BallReflections
'	Next



'Flasher Sols
SolCallBack(15) = "SetLamp 115," 'Pfd Top Panel Flashers(3)

SolCallBack(25) = "SetLamp 125," 'Relaunch + "ON" Flashers
SolCallBack(26) = "SetLamp 126," 'Left Side + "SP" Flashers
SolCallBack(27) = "SetLamp 127," 'Right Side + "AC" Flashers
SolCallBack(28) = "SetLamp 128," 'Top Upper Playfield + "ES" Flashers
SolCallBack(29) = "SetLamp 129," 'Playfield Top Panel + "TA" Flashers
SolCallBack(30) = "SetLamp 130," 'Flame + "TI" Flashers
SolCallBack(31) = "SetLamp 131," 'Station Flashers

InitSideWalls
Sub InitSideWalls
	F26w.y = 1004.5
	F26w.x = -5.87

	F15w.y = 0.1
	F15w.x = 478.2149658

	F28w.y = 0.1
	F28w.x = 478.2149658

	F29w.y = 0.1
	F29w.x = 478.2149658

	F27w.y = 1004.5
	F27w.x = 958.82

	GI_Left.y = 1004.5
	GI_Left.x = -5.87

	GI_Back.y = 0.1
	GI_Back.x = 478.2149658

	GI_Right.y = 1004.5
	GI_Right.x = 958.82

	GIG_Left.y = 1004.5
	GIG_Left.x = -5.87

	GIG_Back.y = 0.1
	GIG_Back.x = 478.2149658

	GIG_Right.y = 1004.5
	GIG_Right.x = 958.82
End Sub

dim aGi(112)
aGi(109) = Array(Gi_White, gi_ambient, GiBlue, Gi_Whitep, GI_Left, Gi_Back, Gi_Right, Gi_Upf1, Gi_Upf2, Gi_Upf3, Gi_Upf4, Gi_Sticker, Gi_UPF, Gi_Shooter, Gi_BumperShine)
aGi(110) = Array(Gi_Green, gig_ambient, Gi_GreenP, Gig_Left, Gig_Back, Gig_Right, Gig_Sticker, Gig_UPF, Gig_Shooter, Gig_BumperShine)
aGi(111) = Array(Gi_bumpers)

SolCallback(9) = "GItoggler"		'PF
SolCallback(10) = "GiSelect"	'Green toggle
SolCallback(11) = "GIbumpers"	'Bumpers

dim CodeGreen : CodeGreen = False
Sub GIselect(value)
	CodeGreen = value' : git2.text = value
	if Value Then
		if LampState(109) + LampState(110) > 0 then
			GIon 1
		End If
	else
		if LampState(109) + LampState(110) > 0 then
			GIon 1
		End If
	End If

End Sub

Sub GIbumpers(value)
	if LampState(111) = 0 then
		Setlamp 111, 1
	else
		Setlamp 111, 0
	End If
End Sub

Sub GIToggler(value)
	if LampState(109) + LampState(110) = 0 then
		GIon 1
		Playsound "Fx_Relay_On", 0, LVL(0.1)
	else
		GIon 0
		Playsound "Fx_Relay_Off", 0, LVL(0.1)
	End If
End Sub

Sub GIon(aEnabled)
	Select Case aEnabled
		Case 0
			SetLamp 110, 0
			Setlamp 109, 0
			RefGreen 5
		Case 1
			If CodeGreen then
				SetLamp 110, 1 : RefGreen True
				Setlamp 109, 0
			Else
				SetLamp 110, 0
				SetLamp 109, 1 : RefGreen False
			End If
	End Select
End Sub

Sub TestGI(nr, number)
	dim x
	for each x in aGi(nr)
		x.IntensityScale = number
	next
End Sub


Sub RefGreen(green)
	dim R,G,B : R = ReflectColor(0)  : G = ReflectColor(1) : B = ReflectColor(2)
	if green then R = 0 : G = 255 : B = 0
	if green = 5 then R = 0 : G = 0 : B = 0
	dim x : for each x in ReflectionsGI : x.Color = RGB(R,G,B) : Next
End Sub


Sub UpdateFlashers()
	FadeGI 109	'white
	'nFadeLmc 109, ReflectionsGI
	nFadeLmc 109, ReflectionsGIB
	FadeGI 110	'green
	'nFadeLmc 110, ReflectionsGIG
	FadeGI 111	'bumpers
	FlashM 111, GI_BumpersT
'	nFadeLmc 111, ReflectionsGIB

	FlashC 115, f15 'Pfd Top Panel Flashers(3)
	FlashOBJm 115, f15p, "Flashers_Orange", 5'12 '(nr, object, imgseq, steps)
	Flashm 115, f15w

	FlashC 125, f25							'Relaunch + "ON" Flashers 6		definitely correct
	Flashmc 125, fbg6

	FlashC 126, f26 							'Left Side + "SP" Flashers 1
	FlashOBJm 126, f26p, "Flashers_Orange", 12
	Flashmc 126, f26w
	Flashmc 126, fbg1

	FlashC 127, f27a								 'Right Side + "AC" Flashers 2
	Flashm 127, f27b
	FlashOBJm 127, f27p, "Flashers_Orange", 12
	Flashmc 127, f27w
	Flashmc 125, fbg2

	FlashC 128, f28a 							'Top Upper Playfield + "ES" Flashers 3
	Flashm 128, f28b
	Flashm 128, F28T	'Transmit Light
	FlashOBJm 128, f28p2, "Flashers_Orange", 12
	Flashmc 128, f28w
	Flashmc 128, fbg3


	FlashC 129, f29 							'Playfield Top Panel + "TA" Flashers 4
	Flashm 129, F29T	'Transmit Light
	Flashmc 129, f29w
	Flashmc 129, fbg4
	Flashmc 129, f29a

'	FlashC 130, f30								 'Flame + "TI" Flashers 5
	FlashC 130, fbg5
	FlashC 131, f31								 'Station Flashers
End Sub

Sub FadeGI(nr)
	dim x
    Select Case FadingLevel(nr)
		Case 3
            FadingLevel(nr) = 0 'completely off
            for each x in aGi(nr)
				x.IntensityScale = FlashLevel(nr)
			next
        Case 4 'off
            FlashLevel(nr) = FlashLevel(nr) - (FlashSpeedDown(nr) * CGT)
            If FlashLevel(nr) < FlashMin(nr) Then
                FlashLevel(nr) = FlashMin(nr)
               FadingLevel(nr) = 3 'completely off
            End if
			for each x in aGi(nr)
				x.IntensityScale = FlashLevel(nr)
			next
        Case 5 ' on
            FlashLevel(nr) = FlashLevel(nr) + (FlashSpeedUp(nr) * CGT)
            If FlashLevel(nr) > FlashMax(nr) Then
                FlashLevel(nr) = FlashMax(nr)
                FadingLevel(nr) = 6 'completely on
            End if
            for each x in aGi(nr)
				x.IntensityScale = FlashLevel(nr)
			next
        Case 6 ' on
			FadingLevel(nr) = 1 'completely on
			for each x in aGi(nr)
				x.IntensityScale = FlashLevel(nr)
			next
    End Select
End Sub



Sub UpdateLamps()
	Motor(0)
	UpdateFlashers
	UpdateBallShadow
	FlashC 1, L1 'FadeL
	FlashC 2, L2 'FadeL
	FlashC 3, L3 'FadeL
	FlashC 4, L4 'FadeL
	FlashC 5, L5 'FadeL
	FlashC 6, L6 'FadeL
	FlashC 7, L7 'FadeL
	FlashC 8, L8 'FadeL
	FlashC 9, L9 'FadeL
	FlashC 10, L10
	FlashC 11, L11
	FlashC 12, L12
	FlashC 13, L13
	FlashC 14, L14	'UPF rollovers
	FlashC 15, L15	'UPF rollovers
	FlashC 16, L16	'UPF rollovers
	FlashC 17, L17
	FlashC 18, L18
	FlashC 19, L19
	FlashC 20, L20
	FlashC 21, L21
	FlashC 22, L22
	FlashC 23, L23
	FlashC 24, L24
	FlashC 25, L25
	FlashC 26, L26
	FlashC 27, L27
	FlashC 28, L28
	FlashC 29, L29
	FlashC 30, L30
	FlashC 31, L31
	FlashC 32, L32
	FlashC 33, L33
	FlashC 34, L34
	FlashC 35, L35
	FlashC 36, L36
'	FlashC 37, L37	'Little Shuttle (insert bd)
	FlashC 38, L38
	FlashC 39, L39
	FlashC 40, L40
	FlashC 41, L41	'Space shuttle toy
	FlashC 42, L42
	FlashC 43, L43
	FlashC 44, L44
	FlashC 45, L45
'	FlashC 46, L46	'Williams (insert bd, left)
'	FlashC 47, L47	'Williams (insert bd, mid)
'	FlashC 48, L48	'Williams (insert bd, right)
''	FlashC 49, L49	'Big Flame #1 (insert bd)
'	nFadeL 49, L49	'Big Flame #1 (insert bd) 	'it's on the backglass and it pretty much just constantly blinks
	FlashC 50, L50
	FlashC 51, L51
	FlashC 52, L52
	FlashC 53, L53
	FlashC 54, L54
	FlashC 55, L55
	FlashC 56, L56
'	FlashOBJm 56, Primmy6, "Flashers_Orange", 12 '(nr, object, imgseq, steps)	'debug
'	FlashOBJm 156, Primmy8, "Flashers_Orange", 12 '(nr, object, imgseq, steps)	'debug
'	FlashC 57, L57	'Big Flame #2 (insert bd)
	FlashC 58, L58
	FlashC 59, L59
	FlashC 60, L60
	FlashC 61, L61
	FlashC 62, L62
	FlashC 63, L63
	FlashC 64, L64
End Sub

dim DisplayColor, DisplayColorG
InitDisplay
Sub InitDisplay()
	dim x
	for each x in Display
		x.Color = RGB(1,1,1)
	next
	Displaytimer.enabled = 1
	DisplayColor = dx.Color
	DisplayColorG = dxG.Color
	If DesktopMode then
		for each x in Display
			x.height = 380 + 70
			x.x = x.x - 956 + 28
			x.y = x.y -32
			x.rotx = -42
			x.visible = 1
		next
		for each x in Display2
			x.x = x.x +100
			x.y = x.y + 1
		next
		for each x in Display3
			x.height = 411
			x.x = x.x +37
			x.y = x.y +30
		next
		for each x in Display4
			x.height = 411
			x.x = x.x +115
			x.y = x.y +30
		next
	Else
		Select Case SingleScreenFS
			Case 1
				Displaytimer.enabled = 1
				for each x in Display
					x.visible = 1
				next
			case 0, 2
				Displaytimer.enabled = 0
				for each x in Display
					x.visible = 0
				next
		End Select
	End If

	if VPXdisplay = 0 Then
		Displaytimer.enabled = 0
		for each x in Display
			x.visible = 0
		next
	End If
End Sub

'Sub tbd_Timer():me.text = Dstate(0,0) & vbnewline & Dstate(0,1) & vbnewline & Dstate(0,2) & _
'				vbnewline & Dstate(0,3) & vbnewline & Dstate(0,4) & vbnewline & Dstate(0,5) & _
'				vbnewline & Dstate(0,6) & vbnewline & Dstate(0,7) & vbnewline & Dstate(0,8) & _
'				vbnewline & Dstate(0,9) & vbnewline & Dstate(0,10) & vbnewline & Dstate(0,11) & _
'				vbnewline & Dstate(0,12) & vbnewline & Dstate(0,13) & vbnewline & Dstate(0,14) & vbnewline & Dstate(0,15)
'End Sub

'num = digits
'chg = LEDs changed since last call
' stat = state

Sub Displaytimer_Timer
	Dim ii, jj, obj, b, x
	Dim ChgLED,num, chg, stat
	ChgLED=Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
		If Not IsEmpty(ChgLED) Then
			For ii=0 To UBound(chgLED)
				num=chgLED(ii, 0) : chg=chgLED(ii, 1) : stat=chgLED(ii, 2)
				For Each obj In Digits(num)
					If chg And 1 Then FadeDisplay obj, stat And 1
					chg=chg\2 : stat=stat\2
				Next
			Next
		End If
End Sub

Sub FadeDisplay(object, onoff)
	If OnOff = 1 Then
		object.color = DisplayColor
	Else
		Object.Color = RGB(1,1,1)
	End If
End Sub

Dim Digits(32)
Digits(0)=Array(D1,D2,D3,D4,D5,D6,D7,D8,D9,D10,D11,D12,D13,D14,D15)
Digits(1)=Array(D16,D17,D18,D19,D20,D21,D22,D23,D24,D25,D26,D27,D28,D29,D30)
Digits(2)=Array(D31,D32,D33,D34,D35,D36,D37,D38,D39,D40,D41,D42,D43,D44,D45)
Digits(3)=Array(D46,D47,D48,D49,D50,D51,D52,D53,D54,D55,D56,D57,D58,D59,D60)
Digits(4)=Array(D61,D62,D63,D64,D65,D66,D67,D68,D69,D70,D71,D72,D73,D74,D75)
Digits(5)=Array(D76,D77,D78,D79,D80,D81,D82,D83,D84,D85,D86,D87,D88,D89,D90)
Digits(6)=Array(D91,D92,D93,D94,D95,D96,D97,D98,D99,D100,D101,D102,D103,D104,D105)
Digits(7)=Array(D106,D107,D108,D109,D110,D111,D112,D113,D114,D115,D116,D117,D118,D119,D120)
Digits(8)=Array(D121,D122,D123,D124,D125,D126,D127,D128,D129,D130,D131,D132,D133,D134,D135)
Digits(9)=Array(D136,D137,D138,D139,D140,D141,D142,D143,D144,D145,D146,D147,D148,D149,D150)
Digits(10)=Array(D151,D152,D153,D154,D155,D156,D157,D158,D159,D160,D161,D162,D163,D164,D165)
Digits(11)=Array(D166,D167,D168,D169,D170,D171,D172,D173,D174,D175,D176,D177,D178,D179,D180)
Digits(12)=Array(D181,D182,D183,D184,D185,D186,D187,D188,D189,D190,D191,D192,D193,D194,D195)
Digits(13)=Array(D196,D197,D198,D199,D200,D201,D202,D203,D204,D205,D206,D207,D208,D209,D210)
Digits(14)=Array(D211,D212,D213,D214,D215,D216,D217,D218)
Digits(15)=Array(D219,D220,D221,D222,D223,D224,D225,D226)
Digits(16)=Array(D227,D228,D229,D230,D231,D232,D233,D234)
Digits(17)=Array(D235,D236,D237,D238,D239,D240,D241,D242)
Digits(18)=Array(D243,D244,D245,D246,D247,D248,D249,D250)
Digits(19)=Array(D251,D252,D253,D254,D255,D256,D257,D258)
Digits(20)=Array(D259,D260,D261,D262,D263,D264,D265,D266)
Digits(21)=Array(D267,D268,D269,D270,D271,D272,D273,D274)
Digits(22)=Array(D275,D276,D277,D278,D279,D280,D281,D282)
Digits(23)=Array(D283,D284,D285,D286,D287,D288,D289,D290)
Digits(24)=Array(D291,D292,D293,D294,D295,D296,D297,D298)
Digits(25)=Array(D299,D300,D301,D302,D303,D304,D305,D306)
Digits(26)=Array(D307,D308,D309,D310,D311,D312,D313,D314)
Digits(27)=Array(D315,D316,D317,D318,D319,D320,D321,D322)


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

'Walls

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

Sub Flashc(nr, object)
    Select Case FadingLevel(nr)
		Case 3
            FadingLevel(nr) = 0 'completely off
            Object.IntensityScale = FlashLevel(nr)
        Case 4 'off
            FlashLevel(nr) = FlashLevel(nr) - (FlashSpeedDown(nr) * CGT)
            If FlashLevel(nr) < FlashMin(nr) Then
                FlashLevel(nr) = FlashMin(nr)
               FadingLevel(nr) = 3 'completely off
            End if
            Object.IntensityScale = FlashLevel(nr)
        Case 5 ' on
            FlashLevel(nr) = FlashLevel(nr) + (FlashSpeedUp(nr) * CGT)
            If FlashLevel(nr) > FlashMax(nr) Then
                FlashLevel(nr) = FlashMax(nr)
                FadingLevel(nr) = 6 'completely on
            End if
            Object.IntensityScale = FlashLevel(nr)
        Case 6 ' on
			FadingLevel(nr) = 1 'completely on
			Object.IntensityScale = FlashLevel(nr)
    End Select
End Sub

Sub FlashObjm(nr, object, imgseq, steps)	'Primitive texture image sequence
	dim x, x2
	Select Case FadingLevel(nr)
		Case 3, 4, 5, 6 'off
			for x = 0 to steps-1
				if FlashLevel(nr) <= ((x/steps) + ((1/steps)/2 )) then
					x2 = x
'					tb3.text = "on " & x & vbnewline & ((x/steps) + ((1/steps)/2 )) & " =? " & x2
					exit for
				end if
			next
			if	FlashLevel(nr) >= 1-(1/steps) then x2 = steps ': tb3.text = "fullon"
		'	if	FlashLevel(nr) <= (1/steps) then x2 = 0 : tb3.text = "fulloff"


'			tb.text = "flashlevel: " & FlashLevel(nr) & vbnewline & "stepper: " & x2 & vbnewline & "flasherimg: " & x2
'			tb.text = FlashLevel(nr) & vbnewline & imgseq & x2
			object.image = imgseq & x2
	End Select
End Sub

Sub Flashm(nr, object) 'multiple flashers, it just sets the flashlevel
	Select Case FadingLevel(nr)
		case 3, 4, 5, 6
			Object.IntensityScale = FlashLevel(nr)
	End Select
End Sub

Sub Flashmc(nr, object) 'multiple flashers conditional
	Select Case FadingLevel(nr)
		case 3, 4, 5, 6
			if HighGraphics = 0 then exit sub
			Object.IntensityScale = FlashLevel(nr)
	End Select
End Sub

Sub NFadeLmc(nr, col) ' used for multiple lights
	dim x
    Select Case FadingLevel(nr)
        Case 3, 4
			if BallReflections = 0 then exit sub
			for each x in col
				x.state = 0
			next
        Case 5, 6
			if BallReflections = 0 then exit sub
			for each x in col
				x.state = 1
			next
    End Select
End Sub






f27a.x = 827.3388
f27a.y = 1340.21
f27a.height = 135

f27b.x = 784.011
f27b.y = 628.534
f27b.height = 121.206
f27b.rotx = -9.1734

f28b.x = 516.72336052
f28b.y = 312.80047786
f28b.height = 111

f28a.x = 101.46988373
f28a.y = 46.54971291
f28a.height = 199


Gi_Ambient.x = 476.470
Gi_Ambient.y = 1004.902
Gig_Ambient.x = 476.470
Gig_Ambient.y = 1004.902

GIblue.x = 395
GIblue.y = 400

GI_Right.y = 1004.5
GI_Right.x = 958.82

GI_Left.y = 1004.5
GI_Left.x = -5.87

GI_Back.y = 0.1
GI_Back.x = 478.2149658

Gi_Sticker.x = 559
Gi_Sticker.y = 183.75
Gig_Sticker.x = 559
Gig_Sticker.y = 183.75

Gi_Upf.x = 290
Gi_Upf.y = 100
GiG_Upf.x = 290
GiG_Upf.y = 100

Gi_Shooter.x = 834.69
Gi_Shooter.y = 1313.123
GiG_Shooter.x = 834.69
GiG_Shooter.y = 1313.123


Gi_BumperShine.x = 357.5
Gi_BumperShine.y = 428
GiG_BumperShine.x = 357.5
GiG_BumperShine.y = 428

InitLamps()


l42.x = 429.7
l42.y = 1578.7
f25.x = 429.7
f25.y = 1578.7

l4.x = 303.5
l4.y = 1419.5

l35.x = 802.25
l35.y = 1107.5

dim lastinput : lastinput = 0
Sub GItype(input)
	dim xg, xp, x, s, rnd
	Select Case input
		case 1	'2700k RGB
			xg = "GI" : xp = "GIP"
			gi_white.opacity = 75000*3
			gi_whiteP.opacity = 2500*3

			gi_left.imagea = "gil"
			gi_back.imagea = "gib"
			gi_right.imagea = "gir"
			gi_left.opacity = 250000 : gi_back.opacity = gi_left.opacity : gi_right.opacity = gi_left.opacity	'38000

			gi_upf.imagea = "gi_lvl2"
			gi_upf.opacity = 1750'1200
			gi_sticker.imagea = "gi_sticker"
			gi_sticker.opacity = 3500'1500
			gi_bumpershine.imagea = "gi_bumpershine"
'			gi_bumpershine.opacity = 1500
			gi_shooter.imagea = "gi_shooter"
'			gi_shooter.opacity = 1500

			ReflectColor(0) = 255
			ReflectColor(1) = 127
			ReflectColor(2) = 0
			for each x in ReflectionsGI : x.Color = RGB(ReflectColor(0),ReflectColor(1),ReflectColor(2) ) : next
'			for x = 0 to (gi2.count -1)
'				On Error Resume Next
'				GI2(x).color = save2700k(x, 0)
'				GI2(x).colorfull = save2700k(x, 1)
'				GI2(x).intensity = save2700k(x, 2)
'			Next
'			for x = 200 to 203		'GI relay on / off	Fading Speeds
'				FlashSpeedUp(x) = 0.01
'				FlashSpeedDown(x) = 0.008
'			Next
'			for x = 300 to 303		'GI	8 step modulation
'				FlashSpeedUp(x) = 0.01
'				FlashSpeedDown(x) = 0.008
'			Next
		case 2	'White
			xg = "GIv" : xp = "GIvP"
			gi_white.opacity = 75000
			gi_whiteP.opacity = 2500'4000

			gi_left.imagea = "gilv"
			gi_back.imagea = "gibv"
			gi_right.imagea = "girv"
'			gi_left.opacity = 38000 : gi_back.opacity = gi_left.opacity : gi_right.opacity = gi_left.opacity
			gi_left.opacity = 150000 : gi_back.opacity = gi_left.opacity : gi_right.opacity = gi_left.opacity

			gi_upf.imagea = "giv_upf"	'oddly named
			gi_upf.opacity = 1200
			gi_sticker.imagea = "giv_sticker"
			gi_sticker.opacity = 1500
			gi_bumpershine.imagea = "giv_bumpershine"
'			gi_bumpershine.opacity = 1500
			gi_shooter.imagea = "giv_shooter"
'			gi_shooter.opacity = 1500

			ReflectColor(0) = 255
			ReflectColor(1) = 255
			ReflectColor(2) = 255
			for each x in ReflectionsGI : x.Color = RGB(ReflectColor(0),ReflectColor(1),ReflectColor(2) ) : next


'			for each x in GI2
'				On Error Resume Next
'				s = mid(x.name, 3, 1)
'				if s = "t" then 'transmit
'					x.Color = RGB(44, 58, 77)
'					x.ColorFull = RGB(188, 188, 155)
'				end If
'			Next
'			git3.colorfull = rgb(255,239,232)	'bulbs
'			git3.color = rgb(253,227,151)
'			git4.colorfull = rgb(255,239,232)
'			git4.color = rgb(253,227,151)
'			git5.colorfull = rgb(255,239,232)
'			git5.color = rgb(253,227,151)
'			for x = 200 to 203		'GI relay on / off
'				FlashSpeedUp(x) = 0.014
'				FlashSpeedDown(x) = 0.014
'			Next
'			for x = 300 to 303		'GI	8 step modulation
'				FlashSpeedUp(x) = 0.014
'				FlashSpeedDown(x) = 0.014
'			Next
		case 0	'Random
			rnd = rndnum(1, 2)
			if rnd <> lastinput then GItype rnd else gitype 0 end if
			Exit Sub
		Case 5 'Sequential
			If LastInput > 1 then LastInput = 0
			gitype (Lastinput+1)
			Exit Sub
		Case Else
			Gitype 0 : exit sub
	End Select
'	dim temp
'	temp = lastinput
	lastinput = input
	GI_white.ImageA = xg
	GI_whitep.ImageA = xp

'	tb.text = temp & " " & input	'debug
End Sub

'Ballshadow routine by Ninuzzu

Dim BallShadow
BallShadow = Array (BallShadow1, BallShadow2, BallShadow3, BallShadow4, BallShadow5, BallShadow6)

Sub UpdateBallShadow()	'called by -1 lamptimer
	On Error Resume Next
    Dim BOT, b
    BOT = GetBalls
	dim CenterPoint : CenterPoint = 425'Table1.Width/2

	' render the shadow for each ball
    For b = 0 to UBound(BOT)
		If BOT(b).X < CenterPoint Then
			BallShadow(b).X = ((BOT(b).X) - (50/6) + ((BOT(b).X - (CenterPoint))/7)) + 10
		Else
			BallShadow(b).X = ((BOT(b).X) + (50/6) + ((BOT(b).X - (CenterPoint))/7)) - 10
		End If

			BallShadow(b).Y = BOT(b).Y + 20
			BallShadow(b).Z = 1
		If BOT(b).Z > 20 Then
			BallShadow(b).visible = 1
		Else
			BallShadow(b).visible = 0
		End If
	Next
End Sub


'cFastFlips by nFozzy
'Bypasses pinmame callback for faster and more responsive flippers
'Version 0.1
'todo replace delay timer with vpmtimer

'Flipper / game-on Solenoid # reference:
'Williams System 11: Sol23 or 24
'Gottlieb System 3: Sol32
'Data East (pre-whitestar): Sol23 or 24
'WPC 90, 92', WPC Security : Sol31

'********************Setup*******************:

'....top of script....
'dim FastFlips

'....init....
'Set FastFlips = new cFastFlips
'with FastFlips
'	.CallBackL = "SolLflipper"	'set flipper sub callbacks
'	.CallBackR = "SolRflipper"	'...
'	.CallBackUL = "SolULflipper"'...(upper flippers, if needed)
'	.CallBackUR = "SolURflipper"'...
'	.TiltObjects = True 'Optional, if True calls vpmnudge.solgameon automatically
'	.Delay = 0			'Optional, if > 0 adds some compensation for solenoid jitter (occasional problem on Bram Stoker's Dracula)
'	.Debug = False		'Debug, always-on flippers. Call FastFlips.Debug True or False in debugger to enable/disable.
'end with

'IF USING DELAY > 0 create a timer object called FastFlipsTimer and add this additional bit of script anywhere:
'Sub FastFlipsTimer_Timer()
'	FastFlips.CutFlippers False
'	me.Enabled = False
'End Sub

'...keydown section... (comment out the upper flippers as needed)
' If KeyCode = LeftFlipperKey then FastFlips.FlipL True :  FastFlips.FlipUL True
' If KeyCode = RightFlipperKey then FastFlips.FlipR True :  FastFlips.FlipUR True
'(Do not use Exit Sub, this script does not handle switch handling at all!)

'...keyUp section...
' If KeyCode = LeftFlipperKey then FastFlips.FlipL False :  FastFlips.FlipUL False
' If KeyCode = RightFlipperKey then FastFlips.FlipR False :  FastFlips.FlipUR False

'...Flipper Callbacks....
'if pinmame flipper callbacks are in use, comment them out (for example 'SolCallback(sLRFlipper)	)
'Use these subs (the ones defined in CallBackL / CallBackR) to handle flipper rotation and sounds

'...Solenoid...
'SolCallBack(31) = "FastFlips.Flippers_On"
'//////for a reference of solenoid numbers, see top /////

'*************************************************

Class cFastFlips
	Public Delay, TiltObjects, Debug
	Private SubL, SubUL, SubR, SubUR, FlippersEnabled

	Private Sub Class_Initialize()
		Delay = 0 : FlippersEnabled = False : Debug = False
	End Sub

	'set callbacks
	Public Property Let CallBackL(aInput)  : Set SubL  = GetRef(aInput) : End Property
	Public Property Let CallBackUL(aInput) : Set SubUL = GetRef(aInput) : End Property
	Public Property Let CallBackR(aInput)  : Set SubR  = GetRef(aInput) : End Property
	Public Property Let CallBackUR(aInput) : Set SubUR = GetRef(aInput) : End Property

	'call callbacks
	Public Sub FlipL(aEnabled)
		if not FlippersEnabled and not debug then Exit Sub
		subL aEnabled
	End Sub

	Public Sub FlipR(aEnabled)
		if not FlippersEnabled and not debug then Exit Sub
		subR aEnabled
	End Sub

	Public Sub FlipUL(aEnabled)
		if not FlippersEnabled and not debug then Exit Sub
		subUL aEnabled
	End Sub

	Public Sub FlipUR(aEnabled)
		if not FlippersEnabled and not debug then Exit Sub
		subUR aEnabled
	End Sub

	Public Sub Flippers_On(aEnabled)	'Handle solenoid
		if delay > 0 and not aEnabled then 	'handle delay
			FastFlipsTimer.Interval = delay
			FastFlipsTimer.Enabled = Not aEnabled
		else
			CutFlippers(aEnabled)
		end if
	End Sub

	Public Sub CutFlippers(aEnabled)
		FlippersEnabled = aEnabled
		if TiltObjects then vpmnudge.solgameon aEnabled
		If Not aEnabled then 'todo delay
			subL False
			subR False
			if not IsEmpty(subUL) then subUL False
			if not IsEmpty(subUR) then subUR False
		End If
	End Sub

End Class


















































'Setup Flipper End Points
Dim EndPointL, EndPointR
EndPointL = 368.761
EndPointR = 492.0299



'==,ggggggggggg,=========================================================================
'dP"""88""""""Y8,              ,dPYb,                                I8
'Yb,  88      `8b              IP'`Yb                                I8
' `"  88      ,8P              I8  8I                         gg  88888888
'     88aaaad8P"               I8  8'                         ""     I8
'     88"""""       ,ggggg,    I8 dP    ,gggg,gg   ,gggggg,   gg     I8    gg     gg
'     88           dP"  "Y8ggg I8dP    dP"  "Y8I   dP""""8I   88     I8    I8     8I
'     88          i8'    ,8I   I8P    i8'    ,8I  ,8'    8I   88    ,I8,   I8,   ,8I
'     88         ,d8,   ,d8'  ,d8b,_ ,d8,   ,d8b,,dP     Y8,_,88,_ ,d88b, ,d8b, ,d8I
'     88         P"Y8888P"    8P'"Y88P"Y8888P"`Y88P      `Y88P""Y888P""Y88P""Y88P"888
'===============================================================================,d8I'====
'               														      ,dP'8I
'Tracks balls, Adjusts ball velocity and adjusts shot angle                  ,8"  8I
'Part 1 - Ball velocity Hack                               	                 I8   8I
'Part 2 - Polarity and Velocity Adjustments on 3 or 5-point envelopes        `8, ,8I
'                                                                             `Y8P"
'Setup -
'Triggers tight to the flippers TriggerLF and TriggerRF. Timers as low as possible, but > 80ms
'Debug box TBpl
'On Flipper Call:
'	ProcessballsL
'	...
'	ProcessballsR
'set up flipper end X coords in variables EndpointL / EndpointR (or in FTS flipper test script)


'-----------Configuration---------------
dim PolarityMod(5, 1), VelMod(5, 1), Ydiminish(3, 1), VelXmod(5, 1)

'Coord PolarityMod, 1, x, y
Sub Coord(N,A,X,Y):a(n, 0) = x :a(n,1) = y : End Sub 'Point #, Array name, XCoord, YCoord

PolarityMod(0,0) = "PolarityMod"
VelMod(0,0) = "VelMod"
Ydiminish(0,0) = "Ydiminish"
VelXmod(0,0) = "VelXmod"

'x = % Position of Flipper
'Y = Output Coefficient for calculation

polarityenabled = True'False

'm1 = 1 : m2 = 1 : m3 = 50 : m4 = 0.93 'Setup Vel Falloff Line
'Coord 1, PolarityMod, 0.38, -4.5		'Early1
'Coord 2, PolarityMod, 0.65, -4		'Early2
'Coord 3, PolarityMod, 0.84, 0		'Middle
'Coord 4, PolarityMod, 0.97, 3.2	'Late3
'Coord 5, PolarityMod, 1.05,  7	'Late4
'
'Coord 1, VelMod, 0.30,	0.9		'Early1
'Coord 2, VelMod, 0.596, 1'0.97		'Early2
'Coord 3, VelMod, 0.782, 1			'Middle
'Coord 4, VelMod, 0.941,  0.95'0.9		'Late3
'Coord 5, VelMod, 1.1, 	0.825'0.85		'Late4

''WPC Steep (71.1/120)
'm1 = 1 : m2 = 1 : m3 = 10 : m4 = 0.935
'Coord 1, PolarityMod, 0.38, -3.5		'Early1
'Coord 2, PolarityMod, 0.596, -5		'Early2
'Coord 3, PolarityMod, 0.8, -2		'Middle
'Coord 4, PolarityMod, 0.97, -1.5	'Late3
'Coord 5, PolarityMod, 1.05,  0	'Late4
''kinda improper. Adjust speed so that these cap out at 1!
'Coord 1, VelMod, 0.30,	0.9		'Early1
'Coord 2, VelMod, 0.596, 0.95'0.97		'Early2
'Coord 3, VelMod, 0.745, 0.965			'Middle
'Coord 4, VelMod, 0.941,  0.95'0.9		'Late3
'Coord 5, VelMod, 1.1, 	0.95'0.85		'Late4

'Sys11 Flat (no polarity adjustment)
m1 = 1 : m2 = 1 : m3 = 50 : m4 = 0.93
Coord 1, PolarityMod, 0.380, 0		'Early1
Coord 2, PolarityMod, 0.596, 0		'Early2
Coord 3, PolarityMod, 0.800, 0		'Middle
Coord 4, PolarityMod, 0.970, 0		'Late3
Coord 5, PolarityMod, 1.050, 0		'Late4

Coord 1, VelMod, 0.30,	1.05		'Early1
Coord 2, VelMod, 0.596, 1'0.97		'Early2
Coord 3, VelMod, 0.782, 1			'Middle
Coord 4, VelMod, 0.941, 1'0.9		'Late3
Coord 5, VelMod, 1.100, 1'0.85		'Late4

'all
Coord 1, Ydiminish, RightFlipper.Y-65,	 0		' Earliest Flipper (keep at 1)
Coord 2, Ydiminish, RightFlipper.Y-11, 	1		' Mid Point
Coord 3, Ydiminish, RightFlipper.Y,	 1	' Latest Flipper


'Part 1 - Overall Velocity Hack
'***************************
Dim LFon, RFon, RF1on, SpeedLimit, M1, M2, M3, M4

Sub TriggerLF_Timer(): LFon = False : me.TimerEnabled = 0 : End Sub
Sub TriggerRF_Timer(): RFon = False : me.TimerEnabled = 0 : End Sub
Sub TriggerRF1_Timer(): RF1on = False : me.TimerEnabled = 0 : End Sub
Sub TriggerRF1_UnHit(): if RF1on then FlipSpeedHack m1, m2, m3, m4, False End If : End Sub

Sub FlipSpeedHack(X1, Y1, X2, Y2, CutoffBool)	'Two points
	if CutoffBool then if activeball.vely > 0 then exit sub	'if ball is going Down, exit sub (inappropriate for upper flippers)
	Dim FinalSpeed : FinalSpeed = BallSpeed(ActiveBall) : if FinalSpeed < x1 then Exit Sub
	Dim VelCoef : VelCoef = SlopeIt(FinalSpeed,X1,Y1,X2,Y2)
		if VelCoef < Y1 then VelCoef = Y1	'Clamp Low
		if VelCoef > Y2 then VelCoef = Y2	'Clamp High
		activeball.velx = activeball.velx * VelCoef
		activeball.vely = activeball.vely * VelCoef

		Dim DebugString : DebugString = "Flip" & vbnewline
		FalloffDebugBox TBflipper, Finalspeed, BallSpeed(ActiveBall), VelCoef, DebugString
End Sub
Sub TBflipper_Timer():me.timerenabled = False : me.text = Empty : End Sub


'=====================================
'Part 2
'Ball Tracking Polarity Correction
'=====================================

'0.09a
'	-	Improved Envelope Functions
'	-	fixed greater than / less than errors
'0.09b - script cleanup, removed unused stuff

Dim Lballstack(9, 5)
Dim Rballstack(9, 5)
'0 = Object reference
'1 = ballID kept in integer (trigger unhit compares this to activeball.ID for wiping ball from stack)
'2 = Ball X pos (set by flip for Polarity correction, wiped on trigger unhit)
'3 = Ball Y pos (TODO)
'4 = Ball X vel
'5 = Partial Flip Coefficient  (kept in 0, 5 only)

Initballstacks
Sub Initballstacks() : dim x: for x = 0 to 9 : Set Lballstack(x,0) = Nothing : Set Rballstack(x,0) = Nothing : next : End Sub

'Left Flipper ====================================

Sub TriggerLF_Hit()	'add a ball to the stack
'	tb.text = activeball.mass
	dim x : for x = 0 to 9
		if Typename(Lballstack(x, 0)) = "Nothing" then
			Set Lballstack(x, 0) = activeball
			Lballstack(x, 1) = activeball.id
			exit For
		End If
	Next
End Sub

Sub TriggerLF_UnHit() 'proc Polarity Correction, then wipe X coords from column 2
	if LFon then 	'FalloffNF
		dim x : for x = 0 to 9	'If X position is set, call Polarity Correction for that object
			if Lballstack(x, 2) > 0 then
				PolarityCorrect Lballstack(x, 0), Lballstack(x, 2), Lballstack(x, 3), Lballstack(x, 4), Lballstack(0, 5), 0
			End If
		Next
		for x = 0 to 9 	'wipe X Positions
			Lballstack(x, 2) = Empty
		Next
		FlipSpeedHack m1, m2, m3, m4, True
	End If
	for x = 0 to 9 	'Remove ball from stack...
		if activeball.id = Lballstack(x, 1) then Set Lballstack(x, 0) = Nothing
	Next
End Sub

Sub ProcessballsL()	'note X position of balls in flipper area
	TriggerLF.TimerEnabled = 1
	LFon = True
	dim x : for x = 0 to 9 'Count X positions of balls in array
		if TypeName(Lballstack(x, 0)) = "IBall" then
			Lballstack(x, 2) = Lballstack(x, 0).X
			Lballstack(x, 3) = Lballstack(x, 0).Y
			Lballstack(x, 4) = Lballstack(x, 0).VelX
		End If
	Next
	'dim totalrotation, currentangler, b
	'CurrentAngler = (LeftFlipper.StartAngle - LeftFlipper.CurrentAngle)
	'TotalRotation = (LeftFlipper.StartAngle - LeftFlipper.EndAngle)
	dim b
	b = ((LeftFlipper.StartAngle - LeftFlipper.CurrentAngle) / (LeftFlipper.StartAngle - LeftFlipper.EndAngle))
	b = abs(b-1)	'invert
	Lballstack(0, 5) = b	'Partial Flip Coefficient
	'tb.text = LeftFlipper.StartAngle - LeftFlipper.EndAngle & vbnewline & leftflipper.CurrentAngle & vbnewline & b
end Sub

'Right Flipper ====================================

Sub TriggerRF_Hit()	'add a ball to the stack
	dim x : for x = 0 to 9
		if Typename(Rballstack(x, 0)) = "Nothing" then
			Set Rballstack(x, 0) = activeball
			Rballstack(x, 1) = activeball.id
			exit For
		End If
	Next
End Sub

Sub TriggerRF_UnHit() 'proc Polarity Correction, then wipe X coords from column 2
	if RFon then 	'FalloffNF
		dim x : for x = 0 to 9	'If X position is set, call Polarity Correction for that object
			if Rballstack(x, 2) > 0 then
				PolarityCorrect Rballstack(x, 0), Rballstack(x, 2), Rballstack(x, 3) , Rballstack(x, 4), Rballstack(0, 5), 1
			End If
		Next
		for x = 0 to 9 	'wipe X Positions
			Rballstack(x, 2) = Empty
		Next
		FlipSpeedHack m1, m2, m3, m4, True
	End If
	for x = 0 to 9 	'Remove ball from stack...
		if activeball.id = Rballstack(x, 1) then Set Rballstack(x, 0) = Nothing
	Next
End Sub

Sub tbBS_Timer()	'debug textbox
'	on error resume next
	dim y(9), x : for x = 0 to 9
		y(x) = Typename(Rballstack(x, 0))
		if TypeName(Rballstack(x, 0)) = "IBall" then y(x) = y(x) & " " & Rballstack(x, 0).ID
		y(x) = y(x) & " " & Rballstack(x, 2)
	Next
	me.text = "Ball 1: " & y(0) & " " & Rballstack(0,1) & vbnewline & _
			  "Ball 2: " & y(1) & " " & Rballstack(1,1) & vbnewline & _
			  "Ball 3: " & y(2) & " " & Rballstack(2,1) & vbnewline & _
			  "Ball 4: " & y(3) & " " & Rballstack(3,1) & vbnewline & _
			  "Ball 5: " & y(4) & " " & Rballstack(4,1) & vbnewline & _
			  "Ball 6: " & y(5) & " " & Rballstack(5,1) & vbnewline & _
			  "Ball 7: " & y(6) & " " & Rballstack(6,1) & vbnewline & _
			  "Ball 8: " & y(7) & " " & Rballstack(7,1) & vbnewline & _
			  "Ball 9: " & y(8) & " " & Rballstack(8,1) & vbnewline & _
			  "Ball10: " & y(9) & " " & Rballstack(9,1) & vbnewline & _
			  "..."
End Sub

Sub ProcessballsR()	'note X position of balls in flipper area
	TriggerRF.TimerEnabled = 1
	RFon = True
	dim x : for x = 0 to 9 'Count X positions of balls in array
		if TypeName(Rballstack(x, 0)) = "IBall" then
			Rballstack(x, 2) = Rballstack(x, 0).X
			Rballstack(x, 3) = Rballstack(x, 0).Y
			Rballstack(x, 4) = Rballstack(x, 0).VelX
		End If
	Next
	dim b
	b = ((RightFlipper.StartAngle - RightFlipper.CurrentAngle) / (RightFlipper.StartAngle - RightFlipper.EndAngle))
	b = abs(b-1)	'invert
	Rballstack(0, 5) = b	'Partial Flip Coefficient
	'tb.text = RightFlipper.StartAngle - RightFlipper.EndAngle & vbnewline & RightFlipper.CurrentAngle & vbnewline & b
end Sub


'Puts an input X through a five-point, four line envelope with flat clamping at the ends
'Function Procedures: Input X, 2D array, special (if True, inverts first two lines for polarity script)
'This 2d Array should start at 1 and end at 5.
'0,0 is used to hold an identifying string for debug purposes
'This 2d Array should hold X data in (x, 0) & Y data in (x, 1)
'Input X, Output Y
Function FivePointEnvelope(xInput, yArray)', special)
	dim y, testF
	If xInput < yArray(2,0) Then	'Setup X Points 	'please keep array X coords sequential!
		y = SlopeIt(xInput, yArray(1,0), yArray(1,1), yArray(2,0), yArray(2, 1)	)
		If yArray(1,1) > yArray(2,1) then
			if y > yArray(1,1) then y = yArray(1,1) 'Clamp Low End
		elseif yArray(1,1) <= yArray(2,1) then
			if y <= yArray(1,1) then y = yArray(1,1) 'Clamp Low End
		End If
		testF = yArray(0,0) & " L1 (early1) " & "x= " & xInput & " y= " & y
	elseif xInput < yArray(3,0) Then	'l2
		y = SlopeIt(xInput, yArray(2,0), yArray(2,1), yArray(3,0), yArray(3,1)	)
		testF =  yArray(0,0) & " L2 (early2)" & "x= " & xInput & " y= " & y
	Elseif xInput < yArray(4,0) Then	'l3
		y = SlopeIt(xInput, yArray(3,0), yArray(3,1), yArray(4,0), yArray(4,1)	)
		testF = yArray(0,0) & " L3 (late3)" & "x= " & xInput & " y= " & y
	Elseif xInput >= yArray(4,0) Then	'l4
		y = SlopeIt(xInput, yArray(4,0), yArray(4,1), yArray(5,0), yArray(5,1)	)
		If yArray(5,1) > yArray(4,1) then 	 'Clamp High End
			if y > yArray(5,1) then y = yArray(5,1)
		elseif yArray(5,1) <= yArray(4,1) then
			if y <= yArray(5,1) then y = yArray(5,1)
		End If
		testF = yArray(0,0) & " L4 (late4) " & " x= " & xInput & " y= " & y
	Else
	debug.print "5error: " & yArray(0,0) & ", xinput = " & xInput : y = 1
	End If
	FivePointEnvelope = y
End Function

dim TestSpecial

Function ThreePointEnvelope(xInput, yArray)
	dim y, test
	If xInput < yArray(2,0) Then
		y = SlopeIt(xInput, yArray(1,0), yArray(1,1), yArray(2,0), yArray(2,1)	)
		If yArray(1,1) > yArray(2,1) then  'Clamp Low End
			if y > yArray(1,1) then y = yArray(1,1)
		elseif yArray(1,1) <= yArray(2,1) then
			if y < yArray(1,1) then y = yArray(1,1)
		End If
		test = "L1 (earliest) " & yArray(0,0) & " x= " & xInput & " y= " & y
	Elseif xInput >= yArray(2,0) Then	'l2
		y = SlopeIt(xInput, yArray(2,0), yArray(2,1), yArray(3,0), yArray(3,1)	)
		If yArray(3,1) > yArray(2,1) then 	 'Clamp High End
			if y > yArray(3,1) then y = yArray(3,1)
		elseif yArray(3,1) <= yArray(2,1) then
			if y < yArray(3,1) then y = yArray(3,1)
		End If
		test = "L3 (latest) " & yArray(0,0) & " x= " & xInput & " y= " & y
	Else
	'debug.print "3error: " & yArray(0,0) & ", xinput = " & xInput
	y = 1
	End If
	'debug.print test
	ThreePointEnvelope = y
End Function

dim PolarityEnabled : PolarityEnabled = True	'debug
Sub PolarityCorrect(object, xpos, ypos, xvel, PartialFLipCoef, LR)	'Corrects angle/velocity using ball data captured at flip
	if TypeName(object) = "Nothing" then Exit Sub 'Bug - This happens when the ball wavers in and out of trigger maybe
	if object.vely > 0 then TBpl.text = "exit sub" : exit sub
	dim TestVar : 	TestVar = "Cutoff"	'debug string
	dim lrcoef : if lr = 1 then lrcoef = -1 else lrcoef = 1 end if	'Direction Coef- could be used to compress the script. readability tho

	dim Y 	 'output (% ball-on-flipper position, 0=base 1=tip)
	if xpos <0.15 then TBpl.text = "xpos<0.15, exit sub " & vbnewline & " y =" & round(y,3) : exit sub 'Cutoff super early
	Select Case LR	'return position of ball on flipper as a % (0=base, 1=tip)
		case 0 : y = SlopeIt(xpos, LeftFlipper.X, 0, EndpointL, 1) 	'base flipper -> 0
		case 1 : y = SlopeIt(xpos, RightFlipper.X, 0, EndpointR, 1)	'End flipper -> 1
	End Select
	if y > 1.05 then y = 1.05 	'Clamp high End


	''''''''''''''declare Polarity Correction + safeties'''''''''''''''
	dim AddX	'Polarity correction
	dim Ycoef:Ycoef = 1	'Safety coef #1 - Cut down Correction if the ball is sufficiently above the flipper base
	if Y > 0.65 then ycoef = ThreePointEnvelope(ypos, Ydiminish)	'Calculate Safety coef #1- if ball is above the flipper
	'PartialFLipCoef -  'Safety coef #2 - handled by processballs, another safety coefficient


	if not PolarityEnabled then 'If Disabled, Exit Sub Here
		TBpl.text = "%" & round(y,5) & vbnewline & "PolarityEnabled = " & PolarityEnabled
		Exit Sub
	End If

	''''''''''''''''''''Apply Velocity Correction''''''''''''''
	dim velcoef 'Overall Velocity coefficient
	Velcoef = FivePointEnvelope(y, VelMod)	'five point velocity envelope based on Y (% on flipper)
	Object.Velx = Object.VelX * velcoef
	Object.Vely = Object.VelY * velcoef

	''''''''''''''''''''Apply Polarity Correction''''''''''''''
	AddX = FivePointEnvelope(y, PolarityMod)*lrcoef	'AddX - Find polarity correction
	object.VelX = object.VelX + 1*(AddX*ycoef*PartialFlipcoef)	'gogo

	'''''''''''''''''''''''Debug Strings''''''''''''''''''''''
	Select Case LR
		case 0 : TestVar =  "Left:" & round(1*(AddX*ycoef*PartialFlipcoef),3) 'debug string  'left flipper
		case 1 : TestVar =  "Right:" & round(1*(AddX*ycoef*PartialFlipcoef),3) 'debug string  'Right Flipper
	End Select
	'debug stuff
	dim d1, d2, d3
	if ycoef < 1 then d1 = "ycoef: " & round(ycoef,5) & vbnewline
	if y = 1.05 then d2 = "(MAX)"
	if PartialFlipcoef < 1 then d3 = "PartialFlipcoef: " & round(PartialFlipcoef,4) & vbnewline
	TBpl.text =  "%" & round(y,3) & d2 & vbnewline & _
				 TestVar & vbnewline & _
				 d1 & d3 & vbnewline & _
				 " "
End Sub

Function BallSpeed(ball) 'Calculates the ball speed
    BallSpeed = SQR(ball.VelX^2 + ball.VelY^2 + ball.VelZ^2)
End Function

'======================================
Function SlopeIt(Input, X1, Y1, X2, Y2)	'Set up line via two points, no clamping. Input X, output Y
	dim x, y, b, m
	x = input
	m = (Y2 - Y1) / (X2 - X1)
	b = Y2 - m*X2

	Y = M*x+b
	SlopeIt = Y
End Function
'======================================

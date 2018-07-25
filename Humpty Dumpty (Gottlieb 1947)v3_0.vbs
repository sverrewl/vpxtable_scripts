' Humpty Dumpty (Gottlieb 1947)

'DOF mapping

'E101 Left flipper
'E102 Right flipper
'E103 Slingshot left
'E104 Slingshot right
'E105 Bumper back left
'E106 Bumper back center
'E107 Bumper back right
'E108 Bumper Left
'E109 Bumper center
'E110 Bumper right
'E111 Knocker
'E112 Shaker
'E113 Gear motor
'E114 Red flashers
'E115 Green flashers
'E116 Blue flashers
'E117 Beacons
'E118 Fan
'E119 Strobes
'E120 Red undercab
'E122 Green undercab
'E123 Blue undercab
'E124 Bell
'E125 Hell ball motor
'E126 Hell ball Green
'E127 Hell ball Red
'E128 Hell ball Blue

'Dim Controller,Tilt,Ball,Balls,Credits,ShootBall,Obj,Bonus,XAdv,Bonuscounter,MultiplierCounter,Tens,Ones,Score,Scoretoadd,TimesDroppedA,TimesDroppedB
'*DOF method for non rom controller tables by Arngrim****************
'*******Use DOF 1**, 1 to activate a ledwiz output*******************
'*******Use DOF 1**, 0 to deactivate a ledwiz output*****************
'*******Use DOF 1**, 2 to pulse a ledwiz output**********************

' Thalamus 2018-07-23
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' No special SSF tweaks yet.

Sub DOF(dofevent, dofstate)
  If B2SOn=True Then
If dofstate = 2 Then
Controller.B2SSetData dofevent, 1:Controller.B2SSetData dofevent, 0
Else  Controller.B2SSetData dofevent, dofstate
End If
End If
End Sub
'********************************************************************
'Sub Table1_Init()
	B2SOn=True

if B2SOn then
		Set Controller = CreateObject("B2S.Server")
		Controller.B2SName = "humpty_dumpty"
		Controller.Run()
		If Err Then MsgBox "Can't Load B2S.Server."
	end if
	 Ball=0
     Balls=5
	 Controller.B2SSetBallInPlay ball
     Controller.B2SSetGameOver 0
Controller.B2SSetmatch 0
     Controller.B2SSetShootAgain 0
 If  Credits > 0 Then
     Controller.B2SSetPlayerup 1
     Controller.B2SSetCanPlay 0
 End If
 If  Credits = 0 Then
     Controller.B2SSetCanPlay 1
 End If


LoadVPM "01560000", "Bally.VBS", 3.26
Sub LoadVPM(VPMver, VBSfile, VBSver)
End Sub

Const BallSize = 50

Dim VarHidden, UseVPMDMD
If Table1.ShowDT = true then
	'UseVPMDMD = True
	'VarHidden = 0
	'lockdown.visible = true
	Hunklight2.visible = true
	Hunklight1.visible = true
	Hunklight3.visible = true
	Hunklight4.visible = true
	Hunklight5.visible = true
	Hunklight6.visible = true
	Hunklight7.visible = true
	A10KL.visible = true
	A20KL.visible = true
	A30KL.visible = true
	A40KL.visible = true
	A50KL.visible = true
	A60KL.visible = true
	A70KL.visible = true
	A80KL.visible = true
	A90KL.visible = true
	CredBox.visible = true
    TiltBox.visible = true
	Light2.visible = true

	'leftrail.visible = true
	'rightrail.visible = true
else
	'UseVPMDMD = False
	'VarHidden = 1
	'lockdown.visible = false
	'leftrail.visible = false
	'rightrail.visible = false
	Hunklight1.visible = false
	Hunklight1.visible = false
	Hunklight2.visible = false
	Hunklight3.visible = false
	Hunklight4.visible = false
	Hunklight5.visible = false
	Hunklight6.visible = false
	Hunklight7.visible = false
	A10KL.visible = false
	A20KL.visible = false
	A30KL.visible = false
	A40KL.visible = false
	A50KL.visible = false
	A60KL.visible = false
	A70KL.visible = false
	CredBox.visible = false
	TiltBox.visible = false
	Light2.visible = false
	A80KL.visible = false
	A90KL.visible = false
	Light10k.visible = false

end if
Dim Score
Dim Ball
Dim Ballsout
Dim Credit
Dim GameOn
Dim Spec1Award
Dim Spec2Award
Dim Spec3Award
Dim Spec4Award
Dim HunKs
Dim TenKs
Dim KBox(9)
Dim HunKLight(7)
Dim BL(10)
Dim Tilted
Dim Tilts
Dim Bonus
Dim Scoring
Dim Hiscore
Dim F
Dim Q
Dim ln
Dim Moo
Dim Hold
Dim BH(5)

dim Ali
dim Baba
dim sinbad
dim patho
dim pathl(16)
dim randy
dim bobo
dim sb
dim bv
dim Amos
dim Otis
dim candy
dim seqo
dim seqi

dim bump1
dim bump2
dim bump3
dim bump4
dim bumpbb1
dim bumpbb2
dim bumpbb3
dim bumpbb4
dim bumptop
dim bumpcenter


Set KBox(1) = A10kL
Set KBox(2) = A20kL
Set KBox(3) = A30kL
Set KBox(4) = A40kL
Set KBox(5) = A50kL
Set KBox(6) = A60kL
Set KBox(7) = A70kL
Set KBox(8) = A80kL
Set KBox(9) = A90kL

Set HunkLight(1) = HunkLight1
Set HunkLight(2) = HunkLight2
Set HunkLight(3) = HunkLight3
Set HunkLight(4) = HunkLight4
Set HunkLight(5) = HunkLight5
Set HunkLight(6) = HunkLight6
Set HunkLight(7) = HunkLight7

Set BL(1) = Light10k
Set BL(2) = Light20k
Set BL(3) = Light40k
Set BL(4) = Light60k
Set BL(5) = Light80k
Set BL(6) = Light100k

Init

Sub Init()
   Score = 0
   Credit = 0
   Randomize
   'LoadHS
LoadTrough.Enabled=1

end sub
'Standard Sounds
   Const SSolenoidOn = "Solenoid"
   Const SSolenoidOff = ""



Sub LoadTrough_Timer()
DOF 103,2
 	trough.CreateBall:trough.Kick 0, 1
trough1.CreateBall:trough1.Kick 0, 1
trough2.CreateBall:trough2.Kick 0, 1
trough3.CreateBall:trough3.Kick 0, 1
trough4.CreateBall:trough4.Kick 0, 1
TroughCount=5
LoadTrough.Enabled=0
 	End Sub

' Key Map
Sub Table1_KeyDown(ByVal keycode)
If PostItHighScoreCheck(keycode) then Exit Sub
	if keycode = 6 then
        if Credit < 26 then
     		Credit = Credit + 1
     		Playsound "coin"
        end if
		CredBox.Text = Credit
	end if
	if keycode = 2 then
        if GameOn = 0 and Credit > 0 then
           StartGame
drop.Rotz = -70
        end if
	end if
	If keycode = PlungerKey Then
		Plunger.PullBack
PlaySound "plungerpull",0,1,0.25,0.25
	End If
    if (keycode =3 or keycode =RightMagnaSave) and gameon = TRUE and BallsOut<= 5 and hold=0 then
	drop.Rotz = 0
DOF 104,2
       BallsOut=BallsOut+1
       PlaySound "balloutr"
       Plungekick.createball.image="JPBall-Dark2"
       Plungekick.kick 270,20
       hold=1

    end if
	If keycode = LeftFlipperKey and tilted = FALSE then
DOF 101,1
		Flipper1.RotateToEnd
		Flipper2.RotateToEnd
		Flipper3.RotateToEnd
		PlaySound "FlipperUp"
	End If
	If keycode = RightFlipperKey and tilted = FALSE then
DOF 102,1
		Flipper4.RotateToEnd
		Flipper5.RotateToEnd
		Flipper6.RotateToEnd
		PlaySound "FlipperUp"
	End If
	If keycode = LeftTiltKey Then
		Nudge 90, 2
		playsound "nudge"
		TiltCheck
	End If
	If keycode = RightTiltKey Then
		Nudge 270, 2
		playsound "nudge"
		TiltCheck
	End If
	If keycode = CenterTiltKey Then
		Nudge 0, 2
		playsound "nudge"
		TiltCheck
	End If
End Sub

Sub Table1_KeyUp(ByVal keycode)
	If keycode = PlungerKey Then
		Plunger.Fire
PlaySound "plunger",0,1,0.25,0.25
	End If
	If keycode = LeftFlipperKey Then
DOF 101,0
		Flipper1.RotateToStart
		Flipper2.RotateToStart
		Flipper3.RotateToStart
		PlaySound "FlipperDown"

	End If
	If keycode = RightFlipperKey Then
DOF 102,0
		Flipper4.RotateToStart
		Flipper5.RotateToStart
		Flipper6.RotateToStart
		PlaySound "FlipperDown"
	End If
End Sub

' Tilt Routine
Sub TiltCheck
   If Tilted = FALSE then
      If Tilts = 0 then
         Tilts = Tilts + int(rnd*100)
         TiltTimer.Enabled = TRUE
      Else
         Tilts = Tilts + int(rnd*120)
      End If
      If Tilts >= 350 and Tilted = FALSE then
         GameOn = FALSE
 		 Flipper1.RotateToStart
 		 Flipper2.RotateToStart
		 Flipper3.RotateToStart
		 Flipper4.RotateToStart
		 Flipper5.RotateToStart
		 Flipper6.RotateToStart
         PlaySound "click"
         Lightb1.State = 0: Primitive110.image = "basedark"
		 gi1.State = 0
         Lightb2.State=  0: Primitive109.image = "basedark"
		 gi2.State = 0
         Lightb3.State = 0: Primitive103.image = "basedark"
		 gi3.State = 0
         Lightb4.State = 0: Primitive105.image = "basedark"
		 gi4.State = 0
         Lightbtopleft.State = 0: Primitive107.image = "basedark"
		 giLightbtopleft.state = 0
         Lightbtopright.State = 0: Primitive108.image = "basedark"
		 giLightbtopright.state = 0
         Lightbbottomleft.State = 0: Primitive104.image = "basedark"
		 giLightbbottomleft.state = 0
         Lightbbottomright.State = 0: Primitive106.image = "basedark"
		 giLightbbottomright.state = 0
		 Light1.State = 0
         LightBottom.State = 0
         RSL1.State = 0: Primitive57.image = "redshootcovertexture"
		 giLighttriangle1.state = 0
DOF 144,0
         RSL2.State = 0: Primitive74.image = "redshootcovertexture"
		 giLighttriangle2.state = 0
DOF 125,0
    	 For X = 1 to 6
       	     BL(x).State = 0
         Next
    	 Tilted = TRUE
         TiltBox.Text = "TILT"

      End If
   End If
End Sub

Sub StartGame()
    Tilted = FALSE
    moo=0
    hold=0
    seqi=0
    Tilts = 0
    PlaySound "reset"
    GameOn = TRUE
DrainWall.IsDropped = TRUE
    Scoring = FALSE
    BallsOut = 0
    If Credit > 0 then
       Credit = Credit - 1
    End if
    CredBox.Text = Credit
    Score = 0
    Hunks = 0
    Tenks = 0
        Spec1Award = FALSE
    Spec2Award = FALSE
    Spec3Award = FALSE
    Spec4Award = FALSE
    lightb1.State = 0: Primitive110.image = "basedark"
	gi1.State = 0
    lightb2.State = 0: Primitive109.image = "basedark"
	gi2.State = 0
    lightb3.State = 0: Primitive103.image = "basedark"
	gi3.State = 0
    lightb4.State = 0: Primitive105.image = "basedark"
	gi4.State = 0
    CRL.state= 1
    crownlight.state= 0
   lefto.state= 0
   righto.state= 0
       Ali=0
    Baba=0
    Amos=0
    Otis=0
    Randy = 1
    bv = 2
    Lightbtopleft.State = 1: Primitive107.image = "baseon"
	giLightbtopleft.State = 1
    Lightbtopright.State = 1: Primitive108.image = "baseon"
	giLightbtopright.state = 1
    Lightbbottomleft.State = 1: Primitive104.image = "baseon"
	giLightbbottomleft.state = 1
    Lightbbottomright.State = 1: Primitive106.image = "baseon"
	giLightbbottomright.state = 1
	'giLightbtop.state = 1: Primitive111.image = "baseon"
	'Lightbtop.state = 1
    Light1.State = 0
    LightBottom.State = 0
    RSL1.State = 0: Primitive57.image = "redshootcovertexture"
	giLighttriangle1.state = 0
DOF 144,0
    RSL2.State = 0: Primitive74.image = "redshootcovertexture"
	giLighttriangle2.state = 0
DOF 125,0
    patho = 1
    For X = 1 to 6
       BL(x).State = 0
    Next
        Ball = 1
    Bonus = 1
    BL(1).State = 1
    For x = 1 to 9
       If x <= 7 then
			Hunklight(x).State = 0
			DOF 200+x, 0
	   End If
       KBox(x).State = 0
	   DOF 10*x, 0
    Next
End Sub

'BallTracking
Dim BallsLOadedForPlay
	BallsLoadedForPlay = 0
'BallTracking
sub drainer_hit()
    drainer.destroyBall
'BallTracking
	BallsLoadedForPlay = BallsLoadedForPlay + 1
	If BallsLoadedForPlay = 5 then
		BallsLoadedForPlay = 0
		DrainWall.IsDropped = 0
	End If
'Ball Tracking
    hold=0
    'BH(Ball).CreateBall
'    ball = Ball + 1
'    if ball = 5 then GameOver
end sub

Sub plungekick_hit()
    Plungekick.kick 270, 5
End Sub

sub GameOver()
msgbox "gameover"
    Ali=0
    Baba=0
    Amos=0
    Otis=0
	MatchTimer.Enabled=1:PlaySound "MotorLeer"
    CRL.state= 0
   crownlight.state= 0
   lefto.state= 0
   righto.state= 0
    If Tilted = FALSE then
If Score = 0 then DrainWall.IsDropped = FALSE
       playsound "motorleer"
       If Score > hiscore then
       hiscore = score
       savehs
       hsbox.text = FormatNumber(hiscore, 0, -1, 0, -1)
       end if
    End If
    GameOn = FALSE
 CheckNewHighScorePostIt1Player hiscore '1 player table
    'CheckNewHighScorePostIt2Player xxx1, xxx2 '2 player table
    'CheckNewHighScorePostIt3Player xxx1, xxx2, xxx3 '3 player table
    'CheckNewHighScorePostIt4Player xxx1, xxx2, xxx3, xxx4 '4 player table

End Sub

' Rotation Sequence Elements


sub tbump1_hit()

	bump1 = 1
	Me.TimerEnabled = 1
	PlaySound "fx_bumper"

	If lightb1.state = 0 then
lightb1.state = 1: addscore (10000)
DOF 148,1
 Primitive110.image = "baseon"
gi1.State = 1
	Ali=Ali+1
	end if
advance()
end sub

sub TBump1_Timer()

           Select Case bump1
               Case 1:skirt1.z = 5:bump1 = 2
               Case 2:skirt1.z = 2:bump1 = 3
               Case 3:skirt1.z = -2:bump1 = 4
               Case 4:skirt1.z = -2:bump1 = 5
               Case 5:skirt1.z = 2:bump1 = 6
               Case 6:skirt1.z = 5:bump1 = 7
               case 7:Me.TimerEnabled = 0
           End Select
DOF 148,0
end sub

sub r2_hit(): PlaySound "fx_bumper"
if rsl2.state= 1 then
Primitive74.image = "redshootcovertexture"
DOF 125,0
giLighttriangle2.state = 0
AwardSpecial
PlaySound "fx_bumper"
end if
if lefto.state= 0 then
addscore (10000)

end if
if lefto.state= 1 then
      if f = 0 then
         f = 5
         Add50KTimer.Enabled = TRUE
      end if
end if
end sub

sub tbump2_hit()

	bump2 = 1
	Me.TimerEnabled = 1
	PlaySound "fx_bumper"

If ali=1 and lightb2.state= 0 then
	lightb2.state = 1
DOF 150,1
addscore (10000)
advance()
Primitive109.image = "baseon"
	gi2.State = 1
Ali=Ali+1
	end if
end sub

sub TBump2_Timer()

           Select Case bump2
               Case 1:skirt2.z = 5:bump2 = 2
               Case 2:skirt2.z = 2:bump2 = 3
               Case 3:skirt2.z = -2:bump2 = 4
               Case 4:skirt2.z = -2:bump2 = 5
               Case 5:skirt2.z = 2:bump2 = 6
               Case 6:skirt2.z = 5:bump2 = 7
               case 7:Me.TimerEnabled = 0
           End Select
DOF 150,0
end sub

sub tbump3_hit()

	bump3 = 1
	Me.TimerEnabled = 1
	PlaySound "fx_bumper"
	If ali=2 and lightb3.state= 0 then
	lightb3.state = 1
DOF 149,1
	addscore (10000)
    advance()
	Primitive103.image = "baseon"
	gi3.State = 1
	Ali=Ali+1


	end if
end sub

sub TBump3_Timer()

           Select Case bump3
               Case 1:skirt3.z = 5:bump3 = 2
               Case 2:skirt3.z = 2:bump3 = 3
               Case 3:skirt3.z = -2:bump3 = 4
               Case 4:skirt3.z = -2:bump3 = 5
               Case 5:skirt3.z = 2:bump3 = 6
               Case 6:skirt3.z = 5:bump3 = 7
               case 7:Me.TimerEnabled = 0
           End Select
DOF 149,0
end sub

sub tbump4_hit()

	bump4 = 1
	Me.TimerEnabled = 1
	PlaySound "fx_bumper"
If ali=3 and lightb4.state = 0 and moo=0 then
	lightb4.state = 1: addscore (10000)
DOF 150,1
	Primitive105.image = "baseon"
	gi4.State = 1
	advance()
	Baba=Baba+1
   playsound "sequence"
   playsound "sequence"
   playsound "sequence"
	light1.state=lightstateon
   lightbottom.state=lightstateon
	moo=1
end if

end sub

sub TBump4_Timer()

           Select Case bump4
               Case 1:skirt4.z = 5:bump4 = 2
               Case 2:skirt4.z = 2:bump4 = 3
               Case 3:skirt4.z = -2:bump4 = 4
               Case 4:skirt4.z = -2:bump4 = 5
               Case 5:skirt4.z = 2:bump4 = 6
               Case 6:skirt4.z = 5:bump4 = 7
               case 7:Me.TimerEnabled = 0
           End Select
DOF 150,0
end sub

sub r5_hit(): PlaySound "fx_bumper"
	if rsl1.state= 1 then
DOF 144,1
	Primitive57.image = "redshootcovertextureon"
	giLighttriangle1.state = 1
	AwardSpecial
	end if
	if righto.state= 0 then
	addscore (10000)
	end if
	if righto.state= 1 then

      if f = 0 then
         f = 5
         Add50KTimer.Enabled = TRUE
      end if
end if
end sub

sub advance()
      If Bonus < 6 then
      If Bonus > 0 then BL(Bonus).State = 0
      Bonus = Bonus + 1
      BL(Bonus).State = 1
               End If
end sub

sub TopBump_hit()

	bumptop = 1
	Me.TimerEnabled = 1
	PlaySound "fx_bumper"
	If Lightbtop.state = 1 then
DOF 136,1
	addscore (10000)
End If
	end sub

sub TopBump_Timer()

           Select Case bumptop
               Case 1:skirttop.z = 5:bumptop = 2
               Case 2:skirttop.z = 2:bumptop = 3
               Case 3:skirttop.z = -2:bumptop = 4
               Case 4:skirttop.z = -2:bumptop = 5
               Case 5:skirttop.z = 2:bumptop = 6
               Case 6:skirttop.z = 5:bumptop = 7
               case 7:Me.TimerEnabled = 0
           End Select
DOF 136,0
end sub

sub CR_hit()
      if f = 0 then
         f = 5
         Add50KTimer.Enabled = TRUE
      end if
   if crownlight.state= 1 then
   AwardSpecial
   crownlight.state= 0
   end if
end sub


sub dbump2_hit()
	bumpcenter = 1
	Me.TimerEnabled = 1
	PlaySound "fx_bumper"
	if dbl2.state= 1 then
DOF 149,1
DOF 151,1
DOF 152,1
DOF 149,1
DOF 151,1
DOF 119,2
DOF 119,2
DOF 119,2
DOF 119,2
pbasecenter.image = "centerbaseon"
giLightcenter.state = 1
      if f = 0 then
         f = 5
         Add50KTimer.Enabled = TRUE
end if
   else
      addscore (10000)

End If

end sub

sub dbump2_Timer()

           Select Case bumpcenter
               Case 1:centerskirt.z = 5:bumpcenter = 2
               Case 2:centerskirt.z = 2:bumpcenter = 3
               Case 3:centerskirt.z = -2:bumpcenter = 4
               Case 4:centerskirt.z = -2:bumpcenter = 5
               Case 5:centerskirt.z = 2:bumpcenter = 6
               Case 6:centerskirt.z = 5:bumpcenter = 7
               case 7:Me.TimerEnabled = 0
           End Select
DOF 149,0
DOF 151,0
DOF 152,0
DOF 119,2
DOF 119,2
DOF 119,2
DOF 119,2
end sub

Dim dw1step, dw2step, dw3step, dw4step, dw5step, dw6step, dw7step, dw8step, dw9step, dw10step, dw11step, dw12step, dw13step, dw14step

'************** Dingwalls

sub dingwall1_hit
	'addscore 1
	rdw1.visible=0
	RDW1a.visible=1
	dw1step=1
	Me.timerenabled=1
end sub

sub dingwall1_timer
	select case dw1step
		Case 1: RDW1a.visible=0: rdw1.visible=1
		case 2:	rdw1.visible=0: rdw1b.visible=1
		Case 3: rdw1b.visible=0: rdw1.visible=1: Me.timerenabled=0
	end Select
	dw1step=dw1step+1
end sub

sub dingwall2_hit
	'addscore 1
	rdw2.visible=0
	rdw3.visible=0
	RDW2a.visible=1
	dw2step=1
	Me.timerenabled=1
end sub

sub dingwall2_timer
	select case dw2step
		Case 1: RDW2a.visible=0: rdw2.visible=1
		case 2:	rdw2.visible=0: rdw2b.visible=1
		Case 3: rdw2b.visible=0: rdw2.visible=1: me.timerenabled=0
	end Select
	dw2step=dw2step+1
end sub

sub dingwall3_hit
	'addscore 1
	Rdw2.visible=0
	Rdw3.visible=0
	RDW3a.visible=1
	dw3step=1
	Me.timerenabled=1
end sub

sub dingwall3_timer
	select case dw3step
		Case 1: RDW3a.visible=0: Rdw2.visible=1
		case 2:	Rdw3.visible=0: rdw3b.visible=1
		Case 3: rdw3b.visible=0: Rdw3.visible=1: me.timerenabled=0
	end Select
	dw3step=dw3step+1
end sub

sub dingwall4_hit
	'addscore 1
	Rdw4.visible=0
	RDW4a.visible=1
	dw4step=1
	Me.timerenabled=1
end sub

sub dingwall4_timer
	select case dw4step
		Case 1: RDW4a.visible=0: Rdw4.visible=1
		case 2:	Rdw4.visible=0: rdw4b.visible=1
		Case 3: rdw4b.visible=0: Rdw4.visible=1: me.timerenabled=0
	end Select
	dw4step=dw4step+1
end sub

sub dingwall5_hit
	'addscore 1
	Rdw5.visible=0
	Rdw6.visible=0
	RDW5a.visible=1
	dw5step=1
	Me.timerenabled=1
end sub

sub dingwall5_timer
	select case dw5step
		Case 1: RDW5a.visible=0: Rdw5.visible=1
		case 2:	Rdw5.visible=0: rdw5b.visible=1
		Case 3: rdw5b.visible=0: Rdw5.visible=1: me.timerenabled=0
	end Select
	dw5step=dw5step+1
end sub

sub dingwall6_hit
	'addscore 1
	Rdw6.visible=0
	Rdw5.visible=0
	RDW6a.visible=1
	dw6step=1
	Me.timerenabled=1
end sub

sub dingwall6_timer
	select case dw6step
		Case 1: RDW6a.visible=0: Rdw6.visible=1
		case 2:	Rdw6.visible=0: rdw6b.visible=1
		Case 3: rdw6b.visible=0: Rdw6.visible=1: me.timerenabled=0
	end Select
	dw6step=dw6step+1
end sub

sub dingwall7_hit
	'addscore 1
	Rdw7.visible=0
	Rdw7.visible=0
	RDW7a.visible=1
	dw7step=1
	Me.timerenabled=1
end sub

sub dingwall7_timer
	select case dw7step
		Case 1: RDW7a.visible=0: Rdw7.visible=1
		case 2:	Rdw7.visible=0: rdw7b.visible=1
		Case 3: rdw7b.visible=0: Rdw7.visible=1: me.timerenabled=0
	end Select
	dw7step=dw7step+1
end sub

sub dingwall8_hit
	'addscore 1
	Rdw8.visible=0
	Rdw7.visible=0
	RDW8a.visible=1
	dw8step=1
	Me.timerenabled=1
end sub

sub dingwall8_timer
	select case dw8step
		Case 1: RDW8a.visible=0: Rdw8.visible=1
		case 2:	Rdw8.visible=0: rdw8b.visible=1
		Case 3: rdw8b.visible=0: Rdw8.visible=1: me.timerenabled=0
	end Select
	dw8step=dw8step+1
end sub


sub dingwall9_hit
	'addscore 1
	Rdw9.visible=0
	Rdw9.visible=0
	RDW9a.visible=1
	dw9step=1
	Me.timerenabled=1
end sub

sub dingwall9_timer
	select case dw9step
		Case 1: RDW9a.visible=0: Rdw9.visible=1
		case 2:	Rdw9.visible=0: rdw9b.visible=1
		Case 3: rdw9b.visible=0: Rdw9.visible=1: me.timerenabled=0
	end Select
	dw9step=dw9step+1
end sub

sub dingwall10_hit
	'addscore 1
	Rdw10.visible=0
	Rdw9.visible=0
	RDW10a.visible=1
	dw10step=1
	Me.timerenabled=1
end sub

sub dingwall10_timer
	select case dw10step
		Case 1: RDW10a.visible=0: Rdw10.visible=1
		case 2:	Rdw10.visible=0: rdw10b.visible=1
		Case 3: rdw10b.visible=0: Rdw10.visible=1: me.timerenabled=0
	end Select
	dw10step=dw10step+1
end sub

sub dingwall11_hit
	'addscore 1
	Rdw11.visible=0
	Rdw11.visible=0
	RDW11a.visible=1
	dw11step=1
	Me.timerenabled=1
end sub

sub dingwall11_timer
	select case dw11step
		Case 1: RDW11a.visible=0: Rdw11.visible=1
		case 2:	Rdw11.visible=0: rdw11b.visible=1
		Case 3: rdw11b.visible=0: Rdw11.visible=1: me.timerenabled=0
	end Select
	dw11step=dw11step+1
end sub

sub dingwall12_hit
	'addscore 1
	Rdw12.visible=0
	Rdw11.visible=0
	RDW12a.visible=1
	dw12step=1
	Me.timerenabled=1
end sub

sub dingwall12_timer
	select case dw12step
		Case 1: RDW12a.visible=0: Rdw12.visible=1
		case 2:	Rdw12.visible=0: rdw12b.visible=1
		Case 3: rdw12b.visible=0: Rdw12.visible=1: me.timerenabled=0
	end Select
	dw12step=dw12step+1
end sub

sub dingwall13_hit
	'addscore 1
	Rdw13.visible=0
	Rdw13.visible=0
	RDW13a.visible=1
	dw13step=1
	Me.timerenabled=1
end sub

sub dingwall13_timer
	select case dw13step
		Case 1: RDW13a.visible=0: Rdw13.visible=1
		case 2:	Rdw13.visible=0: rdw13b.visible=1
		Case 3: rdw13b.visible=0: Rdw13.visible=1: me.timerenabled=0
	end Select
	dw13step=dw13step+1
end sub

sub dingwall14_hit
	'addscore 1
	Rdw14.visible=0
	Rdw13.visible=0
	RDW14a.visible=1
	dw14step=1
	Me.timerenabled=1
end sub

sub dingwall14_timer
	select case dw14step
		Case 1: RDW14a.visible=0: Rdw14.visible=1
		case 2:	Rdw14.visible=0: rdw14b.visible=1
		Case 3: rdw14b.visible=0: Rdw14.visible=1: me.timerenabled=0
	end Select
	dw14step=dw14step+1
end sub

'********** Rubber non-scoring wall animations

'********** Rubber non-scoring wall animations

sub Rwall1_hit
	Rw1.visible=0
	RW1a.visible=1
	rw1step=1
	Me.timerenabled=1
end sub

sub Rwall1_timer
	select case rw1step
		Case 1: RW1a.visible=0: Rw1.visible=1
		case 2:	RW1.visible=0: rw1b.visible=1
		Case 3: rw1b.visible=0: Rw1.visible=1: me.timerenabled=0
	end Select
	rw1step=rw1step+1
end sub

sub Rwall2_hit
	Rw2.visible=0
	RW2a.visible=1
	rw2step=1
	Me.timerenabled=1
end sub

sub Rwall2_timer
	select case rw2step
		Case 1: RW2a.visible=0: Rw2.visible=1
		case 2:	RW2.visible=0: rw2b.visible=1
		Case 3: rw2b.visible=0: Rw2.visible=1: me.timerenabled=0
	end Select
	rw2step=rw2step+1
end sub

sub Rwall1_hit
	Rw1.visible=0
	RW1a.visible=1
	rw1step=1
	Me.timerenabled=1
end sub

sub Rwall1_timer
	select case rw1step
		Case 1: RW1a.visible=0: Rw1.visible=1
		case 2:	RW1.visible=0: rw1b.visible=1
		Case 3: rw1b.visible=0: Rw1.visible=1: me.timerenabled=0
	end Select
	rw1step=rw1step+1
end sub

sub Rwall2_hit
	Rw2.visible=0
	RW2a.visible=1
	rw2step=1
	Me.timerenabled=1
end sub

sub Rwall2_timer
	select case rw2step
		Case 1: RW2a.visible=0: Rw2.visible=1
		case 2:	RW2.visible=0: rw2b.visible=1
		Case 3: rw2b.visible=0: Rw2.visible=1: me.timerenabled=0
	end Select
	rw2step=rw2step+1
end sub





' Bonus Bumpers score up to 30,000


sub bb1_hit()

	bumpbb1 = 1
	Me.TimerEnabled = 1
	PlaySound "fx_bumper"
	If Lightbtopleft.state = 1 then
DOF 129,1
	Primitive107.image = "basedark"
	giLightbtopleft.state = 0
	Lightbtopleft.state = 0
	sb = 1
    bonusbumptimer.enabled = true
	End If
end sub

sub bb1_Timer()

           Select Case bumpbb1
               Case 1:skirtbb1.z = 5:bumpbb1 = 2
               Case 2:skirtbb1.z = 2:bumpbb1 = 3
               Case 3:skirtbb1.z = -2:bumpbb1 = 4
               Case 4:skirtbb1.z = -2:bumpbb1 = 5
               Case 5:skirtbb1.z = 2:bumpbb1 = 6
               Case 6:skirtbb1.z = 5:bumpbb1 = 7
               case 7:Me.TimerEnabled = 0
           End Select
DOF 129,0
end sub


sub bb2_hit()

	bumpbb2 = 1
	Me.TimerEnabled = 1
	PlaySound "fx_bumper"
	If Lightbtopright.state = 1 then
DOF 139,1
	Primitive108.image = "basedark"
	giLightbtopright.state = 0
	Lightbtopright.state = 0
	sb = 1
    bonusbumptimer.enabled = true
End If
end sub

sub bb2_Timer()

           Select Case bumpbb2
               Case 1:skirtbb2.z = 5:bumpbb2 = 2
               Case 2:skirtbb2.z = 2:bumpbb2 = 3
               Case 3:skirtbb2.z = -2:bumpbb2 = 4
               Case 4:skirtbb2.z = -2:bumpbb2 = 5
               Case 5:skirtbb2.z = 2:bumpbb2 = 6
               Case 6:skirtbb2.z = 5:bumpbb2 = 7
               case 7:Me.TimerEnabled = 0
           End Select
DOF 139,0
end sub

sub bb3_hit()

	bumpbb3 = 1
	Me.TimerEnabled = 1
	PlaySound "fx_bumper"
	If Lightbbottomleft.state = 1 then
DOF 132,1
	Primitive104.image = "basedark"
	giLightbbottomleft.state = 0
	Lightbbottomleft.state = 0
	sb = 1
    bonusbumptimer.enabled = true
	End If
end sub

sub bb3_Timer()

           Select Case bumpbb3
               Case 1:skirtbb3.z = 5:bumpbb3 = 2
               Case 2:skirtbb3.z = 2:bumpbb3 = 3
               Case 3:skirtbb3.z = -2:bumpbb3 = 4
               Case 4:skirtbb3.z = -2:bumpbb3 = 5
               Case 5:skirtbb3.z = 2:bumpbb3 = 6
               Case 6:skirtbb3.z = 5:bumpbb3 = 7
               case 7:Me.TimerEnabled = 0
           End Select
DOF 132,0
end sub

sub bb4_hit()

	bumpbb4 = 1
	Me.TimerEnabled = 1
	PlaySound "fx_bumper"
	If Lightbbottomright.state = 1 then
DOF 142,1
	Primitive106.image = "basedark"
	giLightbbottomright.state = 0
	Lightbbottomright.state = 0
	sb = 1
    bonusbumptimer.enabled = true

End If
end sub

sub bb4_Timer()

           Select Case bumpbb4
               Case 1:skirtbb4.z = 5:bumpbb4 = 2
               Case 2:skirtbb4.z = 2:bumpbb4 = 3
               Case 3:skirtbb4.z = -2:bumpbb4 = 4
               Case 4:skirtbb4.z = -2:bumpbb4 = 5
               Case 5:skirtbb4.z = 2:bumpbb4 = 6
               Case 6:skirtbb4.z = 5:bumpbb4 = 7
               case 7:Me.TimerEnabled = 0
           End Select
DOF 142,0
end sub

sub bonusbumptimer_timer()
    If sb > 0 then
       playsound "score"
       changescore
       sb = sb - 1
    else
       bonusbumptimer.enabled = false
    end if
end sub

sub AdvButton_hit()
    playsound "metal"
    Addscore 10000
    wall7x.isdropped=True
	timer2.enabled=True
	timer2.interval=600
    advance()
end sub

sub Timer2_Timer()
timer2.enabled=False
wall7x.isdropped=False
End Sub

dim kick

sub kicker1_hit(): kick = 1: Me.TimerEnabled = 1

   if crownlight.state= 0 then
   playsound "light"
   playsound "light"
   crownlight.state= 1
   lefto.state= 1
   righto.state= 1
   end if
   if light1.state= 1 then
   seqi=6
   end if
   topbonustimer.enabled = TRUE

end sub

sub kicker1_Timer()

           Select Case kick
               Case 1:kick1.z = -43:kick = 2
			   Case 2:kick1.objrotx = 2:kick = 3
               Case 3:kick1.z = -39:kick = 4
			   Case 4:kick1.objrotx = 4:kick = 5
               Case 5:kick1.z = -33:kick = 6
			   Case 6:kick1.objrotx = 6:kick = 7
               Case 7:kick1.z = -33:kick = 8
			   Case 8:kick1.objrotx = 4:kick = 9
               Case 9:kick1.z = -39:kick = 10
			   Case 10:kick1.objrotx = 2:kick = 11
               Case 11:kick1.z = -43:kick = 12
               case 12:Me.TimerEnabled = 0
           End Select
end sub

sub topbonustimer_timer()
       topbonustimer.enabled = false
       seqi=seqi-1
       if Seqi>0 then
       AwardSpecial
       topbonustimer.enabled=True
       end if
       if Seqi<1 then
       if moo=1 then
       light1.state= 0
       end if
       Randomize
       Candy = Int(Rnd(1)* 2) + 1
       If Candy=1 then
       kicker1.kick 123, 5
DOF 106,2
DOF 119,2
DOF 119,2

       end if
       if Candy=2 then
       kicker1.kick 233, 5
DOF 106,2
DOF 119,2
DOF 119,2
       end if
       playsound "saucer"
    end if
end sub

dim kick3

sub kicker2_hit(): kick3 = 1: Me.TimerEnabled = 1
   if lightbottom.state= 1 then AwardSpecial
   bottombonustimer.enabled = TRUE
end sub

sub kicker2_Timer()

           Select Case kick3
               Case 1:kick2.z = -43:kick3 = 2
			   Case 2:kick2.objrotx = 2:kick3 = 3
               Case 3:kick2.z = -39:kick3 = 4
			   Case 4:kick2.objrotx = 4:kick3 = 5
               Case 5:kick2.z = -33:kick3 = 6
			   Case 6:kick2.objrotx = 6:kick3 = 7
               Case 7:kick2.z = -33:kick3 = 8
			   Case 8:kick2.objrotx = 4:kick3 = 9
               Case 9:kick2.z = -39:kick3 = 10
			   Case 10:kick2.objrotx = 2:kick3 = 11
               Case 11:kick2.z = -43:kick3 = 12
               case 12:Me.TimerEnabled = 0
           End Select
end sub


sub bottombonustimer_timer()
    if bonus > 0 then
       PlaySound "bump"
       ChangeScore
    if bonus >1 then
       playsound "bump"
       ChangeScore
    end if
       bl(bonus).State = 0
       bonus = bonus - 1
       If Bonus > 0 then bl(bonus).State = 1
    else
       bottombonustimer.enabled = false
       Bonus = 1
       BL(1).State = 1
       kicker2.kick 180, 5
DOF 109,2
DOF 119,2
DOF 119,2
       playsound "saucer"
    end if
end sub

Sub AddScore(points)
 If Tilted = TRUE then Exit Sub
 If Tilted = FALSE and Scoring = FALSE then
If Score = 0 then DrainWall.IsDropped = FALSE
    ChangeScore
  	Scoring = TRUE
    PlaySound "score"



	Lightbtopleft.state = 0: Primitive107.image = "basedark"
	giLightbtopleft.state = 0


	giLightbtopright.state = 0: Primitive108.image = "basedark"
	Lightbtopright.state = 0

	Lightbbottomleft.state = 0: Primitive104.image = "basedark"
	giLightbbottomleft.state = 0

	Lightbbottomright.state = 0: Primitive106.image = "basedark"
	giLightbbottomright.state = 0


  	Lightbtop.State = 0: Primitive111.image = "basedark"
	giLightbtop.state = 0
  	ScoreTimer.Enabled = TRUE
 End If
End Sub

Sub Add50KTimer_Timer()
    If f > 0 then
       If f = 5 then
          PlaySound "50k"


	Lightbtopleft.state = 0: Primitive107.image = "basedark"
	giLightbtopleft.state = 0


	giLightbtopright.state = 0: Primitive108.image = "basedark"
	Lightbtopright.state = 0

	Lightbbottomleft.state = 0: Primitive104.image = "basedark"
	giLightbbottomleft.state = 0

	Lightbbottomright.state = 0: Primitive106.image = "basedark"
	giLightbbottomright.state = 0



  	    Lightbtop.State = 0: Primitive111.image = "basedark"
		giLightbtop.state = 0

       End If
       ChangeScore
       f = f - 1
    Else
       Add50kTimer.Enabled = FALSE
       Scoring = FALSE


	Lightbtopleft.state = 1: Primitive107.image = "baseon"
	giLightbtopleft.state = 1


	giLightbtopright.state = 1: Primitive108.image = "baseon"
	Lightbtopright.state = 1

	Lightbbottomleft.state = 1: Primitive104.image = "baseon"
	giLightbbottomleft.state = 1

	Lightbbottomright.state = 1: Primitive106.image = "baseon"
	giLightbbottomright.state = 1



  	 Lightbtop.State = 1: Primitive111.image = "baseon"
	 giLightbtop.state = 1
     DBL2.State = 0
	 pbasecenter.image = "centerbasedark"
	 giLightcenter.state = 0
    End If
End Sub


Sub Timer3_Timer()
    If Q > 0 then
       If Q = 30 then
       End If
       ChangeScore
       Q = Q - 1
    Else
       Timer3.Enabled = FALSE
       Scoring = FALSE
    End If
End Sub


Sub ScoreTimer_Timer()
    ScoreTimer.Enabled = FALSE
    Scoring = FALSE
    If Tilted = TRUE then Exit Sub


	Lightbtopleft.state = 1: Primitive107.image = "baseon"
	giLightbtopleft.state = 1


	giLightbtopright.state = 1: Primitive108.image = "baseon"
	Lightbtopright.state = 1

	Lightbbottomleft.state = 1: Primitive104.image = "baseon"
	giLightbbottomleft.state = 1

	Lightbbottomright.state = 1: Primitive106.image = "baseon"
	giLightbbottomright.state = 1



  	Lightbtop.State = 1: Primitive111.image = "baseon"
	giLightbtop.state = 1
    Randomize
    Candy = Int(Rnd(1)* 12) + 1
       If candy<8 then
       DBL2.State= 1
		pbasecenter.image = "centerbaseon"
		giLightcenter.state = 1
       rsl2.state= 0: Primitive74.image = "redshootcovertexture"
		giLighttriangle2.state = 0
DOF 125,0
       rsl1.state= 0: Primitive57.image = "redshootcovertexture"
		giLighttriangle1.state = 0
DOF 144,0
       end if
       If candy=8 then
       DBL2.State= 0
		pbasecenter.image = "centerbasedark"
		giLightcenter.state = 0
       rsl2.state= 0: Primitive74.image = "redshootcovertexture"
		giLighttriangle2.state = 0
DOF 125,0
       rsl1.state= 0: Primitive57.image = "redshootcovertexture"
		giLighttriangle1.state = 0
DOF 144,0
       end if
              If candy=9 then
       DBL2.State= 0
		pbasecenter.image = "centerbasedark"
		giLightcenter.state = 0
       rsl2.state= 0: Primitive74.image = "redshootcovertexture"
		giLighttriangle2.state = 0
DOF 125,0
       rsl1.state= 0: Primitive57.image = "redshootcovertexture"
		giLighttriangle1.state = 0
DOF 144,0
       end if
              If candy>9 then
       DBL2.State= 0
		pbasecenter.image = "centerbasedark"
		giLightcenter.state = 0
       rsl2.state= 1: Primitive74.image = "redshootcovertextureon"
DOF 125,1
		giLighttriangle2.state = 1
       rsl1.state= 1: Primitive57.image = "redshootcovertextureon"
DOF 144,1
		giLighttriangle1.state = 1
       end if
End Sub

Sub ChangeScore()
    If Tilted = TRUE then Exit Sub
	Score = Score + 10000
	' Lite-Up Scoring 10k's
    If Score < 790000 then
       TenKs = (Score Mod 100000) / 10000
       For X = 1 to 9
         KBox(x).State = 0
		 DOF 10*x, 0
       Next
       If Tenks<>0 then
			Kbox(Tenks).State = 1
			DOF 10*Tenks, 1
	   End If
   ' Lite Up Scoring 100K's
       HunKs = Int(Score/100000)
       Hunks = Hunks Mod 10
       For X = 1 to 7
          HunkLight(x).State = 0
		  DOF 200+x, 0
       Next
       If Hunks<>0 then
			HunkLight(Hunks).State = 1
			DOF 200+Hunks, 1
	   End IF
     'Else
       'AuxBox.Text = "Score"
       'AuxSc.Text = FormatNumber(score, 0, -1, 0, -1)
     end if
End Sub

Sub AwardSpecial
DOF 111,2
    If Credit < 26 then
       Credit = Credit + 1
       CredBox.Text = Credit
       PlaySound "knock"
  if f = 0 then
         f = 5
         Add50KTimer.Enabled = TRUE
      end if

    end if
End Sub


Sub OutGate_Hit()
	Hold = 0
    PlaySound "gater"
    If Tilted = TRUE  and Ball = Ballsout then
       GameOn = FALSE
       Exit Sub
    End If
    If Tilted = FALSE then
       Ball = Ball + 1
	'MSGBOX BALL & " SOCRE " & score

       If Ball > 5 then
          GameOn = FALSE
          CheckNewHighScorePostIt1Player score '1 player table

          Chubby=0
         If score > hiscore then
            hiscore = score

			'savehs
            'HSBox.text = FormatNumber(hiscore, 0, -1, 0, -1)
         End If
	   end if
    end if
End Sub

Sub timergate_Timer
	primitive95.rotz=gate.CurrentAngle
End Sub

Sub Gate_hit : PlaySound "Gate" : End Sub


' High Score To Date Routines
sub savehs
	' Based on Black's Highscore routines
	Dim FileObj
	Dim ScoreFile
	Set FileObj=CreateObject("Scripting.FileSystemObject")
	If Not FileObj.FolderExists(UserDirectory) then
		Exit Sub
	End if
	Set ScoreFile=FileObj.CreateTextFile(UserDirectory & "hmpty.txt",True)
		ScoreFile.WriteLine hiscore
		ScoreFile.Close
	Set ScoreFile=Nothing
	Set FileObj=Nothing
end sub

sub loadhs
    ' Based on Black's Highscore routines
	Dim FileObj
	Dim ScoreFile
    dim temp1
    Set FileObj=CreateObject("Scripting.FileSystemObject")
	If Not FileObj.FolderExists(UserDirectory) then
		Exit Sub
	End if
	If Not FileObj.FileExists(UserDirectory & "hmpty.txt") then
		Exit Sub
	End if
	Set ScoreFile=FileObj.GetFile(UserDirectory & "hmpty.txt")
	Set TextStr=ScoreFile.OpenAsTextStream(1,0)
		If (TextStr.AtEndOfStream=True) then
			Exit Sub
		End if
		temp1=TextStr.ReadLine
		TextStr.Close
	    hiscore = CDbl(temp1)
	    Set ScoreFile=Nothing
	    Set FileObj=Nothing
End Sub

Sub UpdateFlipperLogo_Timer
    Primitive1.objrotz = Flipper4.CurrentAngle + 299
    Primitive5.objrotz = Flipper1.CurrentAngle + 57
    Primitive2.objrotz = Flipper5.CurrentAngle + 299
    Primitive4.objrotz = Flipper2.CurrentAngle + 57
    Primitive3.objrotz = Flipper6.CurrentAngle + 299
    Primitive6.objrotz = Flipper3.CurrentAngle + 57
End Sub

sub ln1_hit()
    ln = TRUE
end sub


'****************************************
'  High Score Post-It support 1.2
'
'  1. Copy High Score objects from the template backdrop to the desired table's backdrop. Move if needed
'     (Note: In backdrop Options, make sure Enable EMReels is checked.  If you check and additions items are displayed, move them off the table)

'  2. Import the 4 Post-It note image files into the table using the Import Manager

'  3. Copy this High Score Post-It script into the table script

'  4. Change the HighScoreFilename, "xxxxHighscorePostIt.txt" below - must be unique for each table

'  5. Add this line to the top of the KeyDown sub: 	If PostItHighScoreCheck(keycode) then Exit Sub

'  6. Add one of the following lines below to the bottom of the Game Over sub of the table.
'     Note:  Each table is different, you will have to figure out the scoring variable and the game over/match over subroutine
'     Replace xxxx with the variable used to track scoring, usually Score, Score(1), Score1, PlayerScore, etc
'    CheckNewHighScorePostIt1Player xxxx 					'1 player table
'    CheckNewHighScorePostIt2Player xxx1, xxx2 				'2 player table
'    CheckNewHighScorePostIt3Player xxx1, xxx2, xxx3		'3 player table
'    CheckNewHighScorePostIt4Player xxx1, xxx2, xxx3, xxx4	'4 player table
'****************************************
Const HighScoreFilename = "HumptyDumpty.txt"

Dim HSAHighScore, HSA1, HSA2, HSA3
Dim HSScore100k, HSScore10k, HSScoreK, HSScore100, HSScore10, HSScore1, HSScorex	'Define 5 different score values for each reel to use
Dim HSEnterMode, hsLetterFlash, hsEnteredDigits(3), hsCurrentDigit, hsCurrentLetter
Const hsFlashDelay = 4
Const DefaultHighScore = 20000
Const DefaultHSA1 = 78
Const DefaultHSA2 = 74
Const DefaultHSA3 = 78

LoadHighScore
Sub LoadHighScore
	Dim FileObj
	Dim ScoreFile
	Dim TextStr
    Dim SavedDataTemp3 'HighScore
    Dim SavedDataTemp4 'HSA1
    Dim SavedDataTemp5 'HSA2
    Dim SavedDataTemp6 'HSA3
    Set FileObj=CreateObject("Scripting.FileSystemObject")
	If Not FileObj.FolderExists(UserDirectory) then
		Exit Sub
	End if
	If Not FileObj.FileExists(UserDirectory & HighScoreFilename) then
		SetDefaultHSTD:UpdatePostIt:SaveHighScore
		Exit Sub
	End if
	Set ScoreFile=FileObj.GetFile(UserDirectory & HighScoreFilename)
	Set TextStr=ScoreFile.OpenAsTextStream(1,0)
		If (TextStr.AtEndOfStream=True) then
			SetDefaultHSTD:UpdatePostIt:SaveHighScore
			Exit Sub
		End if
		SavedDataTemp3=Textstr.ReadLine ' HighScore
		SavedDataTemp4=Textstr.ReadLine ' HSA1
		SavedDataTemp5=Textstr.ReadLine ' HSA2
		SavedDataTemp6=Textstr.ReadLine ' HSA3
		TextStr.Close
		HSAHighScore=SavedDataTemp3
		HSA1=SavedDataTemp4
		HSA2=SavedDataTemp5
		HSA3=SavedDataTemp6
		UpdatePostIt
	    Set ScoreFile = Nothing
	    Set FileObj = Nothing
End Sub

Sub SetDefaultHSTD  'bad data or missing file - reset and resave
	HSAHighScore = DefaultHighScore
	HSA1 = DefaultHSA1
	HSA2 = DefaultHSA2
	HSA3 = DefaultHSA3
	SaveHighScore
End Sub

Sub LastScoreReels
		HSScorex = LastScore
		HSScore100K=Int (HSScorex/100000)'Calculate the value for the 100,000's digit
		HSScore10K=Int ((HSScorex-(HSScore100k*100000))/10000) 'Calculate the value for the 10,000's digit
		HSScoreK=Int((HSScorex-(HSScore100k*100000)-(HSScore10K*10000))/1000) 'Calculate the value for the 1000's digit
		HSScore100=Int((HSScorex-(HSScore100k*100000)-(HSScore10K*10000)-(HSScoreK*1000))/100) 'Calculate the value for the 100's digit
		HSScore10=Int((HSScorex-(HSScore100k*100000)-(HSScore10K*10000)-(HSScoreK*1000)-(HSScore100*100))/10) 'Calculate the value for the 10's digit
		HSScore1=Int(HSScorex-(HSScore100k*100000)-(HSScore10K*10000)-(HSScoreK*1000)-(HSScore100*100)-(HSScore10*10)) 'Calculate the value for the 1's digit

End Sub

Sub UpdatePostIt
		HSScorex = HSAHighScore
		HSScore100K=Int (HSScorex/100000)'Calculate the value for the 100,000's digit
		HSScore10K=Int ((HSScorex-(HSScore100k*100000))/10000) 'Calculate the value for the 10,000's digit
		HSScoreK=Int((HSScorex-(HSScore100k*100000)-(HSScore10K*10000))/1000) 'Calculate the value for the 1000's digit
		HSScore100=Int((HSScorex-(HSScore100k*100000)-(HSScore10K*10000)-(HSScoreK*1000))/100) 'Calculate the value for the 100's digit
		HSScore10=Int((HSScorex-(HSScore100k*100000)-(HSScore10K*10000)-(HSScoreK*1000)-(HSScore100*100))/10) 'Calculate the value for the 10's digit
		HSScore1=Int(HSScorex-(HSScore100k*100000)-(HSScore10K*10000)-(HSScoreK*1000)-(HSScore100*100)-(HSScore10*10)) 'Calculate the value for the 1's digit
		EMReelHSNum1.SetValue(HSScore100K):If HSScorex<100000 Then EMReelHSNum1.SetValue 10
		EMReelHSNum2.SetValue(HSScore10K):If HSScorex<10000 Then EMReelHSNum2.SetValue 10
		EMReelHSNum3.SetValue(HSScoreK):If HSScorex<1000 Then EMReelHSNum3.SetValue 10:EMReelHSComma.SetValue 0:Else EMReelHSComma.SetValue 1
		EMReelHSNum4.SetValue(HSScore100):If HSScorex<100 Then EMReelHSNum4.SetValue 10
		EMReelHSNum5.SetValue(HSScore10):If HSScorex<10 Then EMReelHSNum5.SetValue 10
		EMReelHSNum6.SetValue(HSScore1)
		EMReelHSName1.SetValue HSA1
		EMReelHSName2.SetValue HSA2
		EMReelHSName3.SetValue HSA3
End Sub

Sub SaveHighScore
	Dim FileObj
	Dim ScoreFile
	Set FileObj=CreateObject("Scripting.FileSystemObject")
	If Not FileObj.FolderExists(UserDirectory) then
		Exit Sub
	End if
	Set ScoreFile=FileObj.CreateTextFile(UserDirectory & HighScoreFilename,True)
		ScoreFile.WriteLine HSAHighScore
		ScoreFile.WriteLine HSA1
		ScoreFile.WriteLine HSA2
		ScoreFile.WriteLine HSA3
		ScoreFile.Close
	Set ScoreFile=Nothing
	Set FileObj=Nothing
End Sub

Sub HighScoreEntryInit()
	HSEnterMode = True
	hsCurrentDigit = 0
	hsCurrentLetter = 1:HSA1=1
	HighScoreFlashTimer.Interval = 250
	HighScoreFlashTimer.Enabled = True
	hsLetterFlash = hsFlashDelay
End Sub

Sub HighScoreFlashTimer_Timer()
	hsLetterFlash = hsLetterFlash-1
	If hsLetterFlash=1 then 'switch to underscore
		Select Case hsCurrentLetter
			Case 1:
				EMReelHSName1.SetValue 28
			Case 2:
				EMReelHSName2.SetValue 28
			Case 3:
				EMReelHSName3.SetValue 28
		End Select
	End If
	If hsLetterFlash=0 then 'switch back
		hsLetterFlash = hsFlashDelay
		Select Case hsCurrentLetter
			Case 1:
				EMReelHSName1.SetValue HSA1
			Case 2:
				EMReelHSName2.SetValue HSA2
			Case 3:
				EMReelHSName3.SetValue HSA3
		End Select
	End If
End Sub

Sub HighScoreProcessKey(keycode)
    If keycode = LeftFlipperKey Then
		hsLetterFlash = hsFlashDelay
		Select Case hsCurrentLetter
			Case 1:
				HSA1=HSA1-1:If HSA1=-1 Then HSA1=26 'no backspace on 1st digit
				UpdatePostIt
			Case 2:
				HSA2=HSA2-1:If HSA2=-1 Then HSA2=27
				UpdatePostIt
			Case 3:
				HSA3=HSA3-1:If HSA3=-1 Then HSA3=27
				UpdatePostIt
		 End Select
    End If

	If keycode = RightFlipperKey Then
		hsLetterFlash = hsFlashDelay
		Select Case hsCurrentLetter
			Case 1:
				HSA1=HSA1+1:If HSA1>27 Then HSA1=0
				UpdatePostIt
			Case 2:
				HSA2=HSA2+1:If HSA2>27 Then HSA2=0
				UpdatePostIt
			Case 3:
				HSA3=HSA3+1:If HSA3>27 Then HSA3=0
				UpdatePostIt
		 End Select
	End If

    If keycode = StartGameKey Then
		Select Case hsCurrentLetter
			Case 1:
				hsCurrentLetter=2 'ok to advance
				HSA2=HSA1 'start at same alphabet spot
				EMReelHSName1.SetValue HSA1:EMReelHSName2.SetValue HSA2
			Case 2:
				If HSA2=27 Then 'bksp
					HSA2=0
					hsCurrentLetter=1
				Else
					hsCurrentLetter=3 'enter it
					HSA3=HSA2 'start at same alphabet spot
				End If
				EMReelHSName2.SetValue HSA2:EMReelHSName3.SetValue HSA3
			Case 3:
				If HSA3=27 Then 'bksp
					HSA3=0
					hsCurrentLetter=2
				Else
					SaveHighScore 'enter it
					HighScoreFlashTimer.Enabled = False
					HSEnterMode = False

					UpdatePostIt
					EMReelHSTitle.SetValue 0
				End If
				EMReelHSName3.SetValue HSA3
		End Select
    End If
End Sub

Function PostItHighScoreCheck (keycode)
	PostItHighScoreCheck = 0
    If HSEnterMode Then
		HighScoreProcessKey(keycode)

		Select Case keycode
			Case LeftFlipperKey, RightFlipperKey, 2, StartGameKey
				PostItHighScoreCheck = 1
		End Select
	End If
End Function

Sub CheckNewHighScorePostIt (newScore)
		If CLng(newScore) > CLng(HSAHighScore) Then
			HSAHighScore=newScore:HSA1 = 0:HSA2 = 0:HSA3 = 0:UpdatePostIt
			EMReelHSTitle.SetValue 1
			HighScoreEntryInit()
		End If
End Sub
Sub CheckNewHighScorePostIt1Player (newScore1)
	CheckNewHighScorePostIt newScore1
End Sub
Sub CheckNewHighScorePostIt2Player (newScore1, newScore2)
	Dim bestscore
	bestscore = newScore1
	If newScore2 > bestscore then bestscore = newScore2
	CheckNewHighScorePostIt bestscore
End Sub
Sub CheckNewHighScorePostIt3Player (newScore1, newScore2, newScore3)
	Dim bestscore
	bestscore = newScore1
	If newScore2 > bestscore then bestscore = newScore2
	If newScore3 > bestscore then bestscore = newScore3
	CheckNewHighScorePostIt bestscore
End Sub
Sub CheckNewHighScorePostIt4Player (newScore1, newScore2, newScore3, newScore4)
	Dim bestscore
	bestscore = newScore1
	If newScore2 > bestscore then bestscore = newScore2
	If newScore3 > bestscore then bestscore = newScore3
	If newScore4 > bestscore then bestscore = newScore4
	CheckNewHighScorePostIt bestscore
End Sub

'************************************
' What you need to add to your table
'************************************

' a timer called CollisionTimer. With a fast interval, like from 1 to 10
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

' The Double For loop: This is a double cycle used to check the collision between a ball and the other ones.
' If you look at the parameters of both cycles, you’ll notice they are designed to avoid checking
' collision twice. For example, I will never check collision between ball 2 and ball 1,
' because I already checked collision between ball 1 and 2. So, if we have 4 balls,
' the collision checks will be: ball 1 with 2, 1 with 3, 1 with 4, 2 with 3, 2 with 4 and 3 with 4.

' Sum first the radius of both balls, and if the height between them is higher then do not calculate anything more,
' since the balls are on different heights so they can't collide.

' The next 3 lines calculates distance between xth and yth balls with the Pytagorean theorem,

' The first "If": Checking if the distance between the two balls is less than the sum of the radius of both balls,
' and both balls are not already colliding.

' Why are we checking if balls are already in collision?
' Because we do not want the sound repeting when two balls are resting closed to each other.

' Set the collision property of both balls to True, and we assign the number of the ball colliding

' Play the collide sound of your choice using the VOL, PAN and PITCH functions

' Last lines: If the distance between 2 balls is more than the radius of a ball,
' then there is no collision and then set the collision property of the ball to False (-1).


Sub Pins_Hit (idx)
	PlaySound "pinhit_low", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Targets_Hit (idx)
	PlaySound "target", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub metal_Hit (idx)
	PlaySound "metal", 10, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Metals_Medium_Hit (idx)
	PlaySound "metalhit_medium", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Metals2_Hit (idx)
	PlaySound "metalhit2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Gates_Hit (idx)
	PlaySound "gate4", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Spinner_Spin
	PlaySound "fx_spinner",0,.25,0,0.25
End Sub

Sub rubbers_Hit(idx)
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 1 then
		PlaySound "fx_rubber2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End if
	If finalspeed >= 6 AND finalspeed <= 1 then
 		RandomSoundRubber()
 	End If
End Sub

Sub Posts_Hit(idx)
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 16 then
		PlaySound "fx_rubber2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End if
	If finalspeed >= 6 AND finalspeed <= 16 then
 		RandomSoundRubber()
 	End If
End Sub

Sub RandomSoundRubber()
	Select Case Int(Rnd*3)+1
		Case 1 : PlaySound "rubber_hit_1", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 2 : PlaySound "rubber_hit_2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 3 : PlaySound "rubber_hit_3", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
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
		Case 1 : PlaySound "flip_hit_1", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 2 : PlaySound "flip_hit_2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 3 : PlaySound "flip_hit_3", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End Select
End Sub

Sub Table1_Exit
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
'    JP's VP10 Collision & Rolling Sounds
'*****************************************

Const tnob = 10 ' total number of balls
ReDim rolling(tnob)
ReDim collision(tnob)
Initcollision

Sub Initcollision
    Dim i
    For i = 0 to tnob
        collision(i) = -1
        rolling(i) = False
    Next
End Sub

Sub CollisionTimer_Timer()
    Dim BOT, B, B1, B2, dx, dy, dz, distance, radii
    BOT = GetBalls

    ' rolling

	For B = UBound(BOT) +1 to tnob
        rolling(b) = False
        StopSound("fx_ballrolling" & b)
	Next

    If UBound(BOT) = -1 Then Exit Sub

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


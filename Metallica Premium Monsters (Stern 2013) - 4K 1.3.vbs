  Option Explicit
    Randomize

' Thalamus 2018-07-24
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' No special SSF tweaks yet.
' Added InitVpmFFlipsSAM


Const VolDiv = 2000

	 Const cGameName = "mtl_170hc"
	 'Const cGameName = "mtl_170h"
     Const UseSolenoids = 1
     Const UseLamps = 0
     Const UseSync = 1
     Const HandleMech = 0
     Const SSolenoidOn = "Solenoid"
     Const SSolenoidOff = ""
     Const SCoin = "CoinIn"
	Const UseVPMModSol = 1


    Dim xx
    Dim Bump1, Bump2, Bump3, Mech3bank,bsTrough,bsLHole, bsRHole,dtl,cbRight,turntable,cbCaptive,cbCaptive2
	Dim PlungerIM,LMag,RMag, HMag, HMag2
	Dim cBall
	Dim mCaptive
	Dim VarHidden, UseVPMDMD
	Dim CrossMech
	Dim DayNight
	Dim SideWallGI
	Dim SparkyMod
	Dim CemMod
	Dim LeftCard, CenterCard, RightCard
	Dim LeftImage, CenterImage, RightImage
	LeftImage=array("apron_left_1", "apron_left_2", "apron_left_3")
	CenterImage=array("apron_center_1", "apron_center_2", "apron_center_3")
	RightImage=array("apron_right_1", "apron_right_2", "apron_right_3", "apron_right_4")

	'mods/OPTIONS
	Dim SnakeSkin, CoffinMod
	CoffinMod = 2 ' 0 = off, 2 = skulls
	SnakeSkin = 1 ' 0 = odd, 1 = on
	SideWallGI = 1 ' 0 = off, 1 = On
	SparkyMod = 1 ' 0 = original, 1 = Modded Paint scheme
	CemMod = 1 ' 0 = No cemetery arch, 1 = cemetery arch
	LeftCard = 0 ' 0 = standard yellow, 1 = standard white, 2 = custom black
	CenterCard = 1 ' 0 = premium, 1 = custom, 2 = pro
	RightCard = 0 ' 0 = standard yellow, 1 = standard white, 2 = custom black, 3 alt custom black

	If table.ShowDT = true then
	'	UseVPMDMD = true
		VarHidden = 0
		lockdown.visible = true
	else
	'	UseVPMDMD = False
		VarHidden = 0
		lockdown.visible = false
	end if

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "Can't open controller.vbs"
On Error Goto 0

  LoadVPM "01560000", "sam.VBS", 3.10


	Set cBall = ckicker.createball
	ckicker.Kick 0, 0

Sub table_Init
'    VPMInit Me
	UpPost.Isdropped=true
	With Controller
		.GameName = cGameName
		If Err Then MsgBox "Can't start Game " & cGameName & vbNewLine & Err.Description:Exit Sub
		.SplashInfoLine = "Metallica (Stern 2013)"
		.HandleKeyboard = 0
		.ShowTitle = 0
		.ShowDMDOnly = 1
		.ShowFrame = 0
		.HandleMechanics = 1
		.Hidden = VarHidden
		On Error Resume Next
		.Run GetPlayerHWnd
		If Err Then MsgBox Err.Description
	End With
	InitVpmFFlipsSAM
    On Error Goto 0

'Trough
    Set bsTrough = New cvpmBallStack
    bsTrough.InitSw 0, 21, 20, 19, 18, 0, 0, 0
    bsTrough.InitKick BallRelease, 90, 8
    bsTrough.InitExitSnd SoundFX("ballrelease",DOFContactors), "Solenoid"
    bsTrough.Balls = 4

'Hole bsRHole - right eject
    Set bsRHole = New cvpmBallStack
      With bsRHole
          .InitSw 0, 51, 0, 0, 0, 0, 0, 0
          .InitKick kicker3, 40, 56
          .InitExitSnd SoundFX("popper_ball",DOFContactors), "Solenoid"
          '.KickForceVar = 2
      End With

	Set mCaptive = New cvpmMagnet : With mCaptive
		.InitMagnet CaptBall, 100
		.CreateEvents "mCaptive"
		.GrabCenter = True
		.MagnetOn = True
	End With

	Fkicker.createball:fkicker.kick 0,0,0
	ekicker.createball:ekicker.kick 0,0,0
	HammerP.RotX = 35

	if SparkyMod = 1 Then
		SparkyBody.image = "Sparky2Map"
		SparkyHead.image = "SparkyMapping2"
	Else
		SparkyBody.image = "SparkyMap3"
		SparkyHead.image = "SparkyMapping3"
	End if

	if CemMod = 1 Then
		CemeteryP.visible = 1
	Else
		CemeteryP.visible = 0
	End If

	DayNight = Table.NightDay 'read day/night slider value
	Intensity ' Sets GI brightness depending on day/night slider settings

	Set LMag=New cvpmMagnet
	LMag.InitMagnet LMagnet,22
	LMag.GrabCenter=True

	Set RMag=New cvpmMagnet
	RMag.InitMagnet RMagnet,40
	RMag.GrabCenter=True

	Set HMag=New cvpmMagnet
	HMag.InitMagnet HMagnet,40
	HMag.GrabCenter=True

	l31f.visible = 0
	l31f2.visible = 0

	If SnakeSkin = 1 Then
		Primitive55.visible = 0
		Primitive118.visible = 1
		Primitive27.material = "GreenWire"
		Primitive28.image = "RRampMap2"
	Else
		Primitive55.visible = 1
		Primitive118.visible = 0
		Primitive27.material = "Metal Black"
		Primitive28.image = ""
	End If

	vpmNudge.TiltSwitch=-7
	vpmNudge.Sensitivity= 1
	vpmNudge.TiltObj=Array(Bumper1b,Bumper2b,Bumper3b,LeftSlingshot,RightSlingshot)

'Apron Cards
	ApronCardCenter.image = CenterImage(CenterCard)
	ApronCardLeft.image = LeftImage(LeftCard)
	ApronCardRight.image = RightImage(RightCard)

End Sub

 Sub table_Paused:Controller.Pause = 1:End Sub
 Sub table_unPaused:Controller.Pause = 0:End Sub

 Sub table_KeyDown(ByVal keycode)
    If keycode = PlungerKey Then Plunger.Pullback
    If vpmKeyDown(keycode) Then Exit Sub
End Sub

Sub table_KeyUp(ByVal keycode)
	If vpmKeyUp(keycode) Then Exit Sub
	If Keycode = StartGameKey Then Controller.Switch(16) = 0
    If keycode = PlungerKey Then
		Plunger.Fire
        PlaySound "Plunger"
	End If
End Sub

SolCallback(1) = "solTrough"
SolCallback(2) = "solAutofire"
SolCallback(3) = "LMag.MagnetOn="
SolCallback(4) = "RMag.MagnetOn="
SolCallback(5) = "SnakeKick"
SolCallback(6) = "ScoopKick"
'SolCallback(7) = "orbitpost"  'premium not used
'SolCallback(8) = shaker motor ' optional
SolCallback(9)  = "vpmSolSound SoundFX(""jet3"",DOFContactors),"
SolCallback(10) = "vpmSolSound SoundFX(""jet3"",DOFContactors),"
SolCallback(11) = "vpmSolSound SoundFX(""jet3"",DOFContactors),"
SolCallback(12) = "SnakeJawLatch"  ' premium snake jaw latch
'SolCallback(13) = Left Slingshot
'SolCallback(14) = Right Slingshot
SolCallback(15) = "SolLFlipper"
SolCallback(16) = "SolRFlipper"
'SolCallback(17) = not used
SolCallback(18) = "SolSparkyHead"
SolModCallback(19) = "Flash19"
SolCallback(20) = "CrossMotor"' premium grave marker motor
SolCallback(21) = "SetLamp 121,"
SolCallback(22) = "SetLamp 122,"
SolModCallback(23) = "Flash23"
SolCallback(25) = "SetLamp 125,"
SolModCallback(26) = "Flash26"
SolModCallback(27) = "Flash27"
SolModCallback(28) = "Flash28"
SolModCallback(29) = "Flash29"
SolModCallback(30) = "Flash30" 'premium snake flasher 'captive ball flasher
SolModCallback(31) = "Flash31" 'coffin insert flasher x2
SolmodCallback(32) = "Flash32" 'electric chair insert flasher
SolCallback(51) = "CLockRelease"
SolCallback(52) = "CoffinMagDown" 'Premium coffinMagnet Down
SolCallback(53) = "Hammer" 'premium hammer
SolCallback(54) = "resetinline" ' premium DTresets
Solcallback(55) = "OrbitPost" ' premium loop post
SolCallback(56) = "SolSnakeJaw" '  premium Snake Jaw
SolCallback(57) = "HMag.MagnetOn=" ' premium CoffinMag On

controller.switch(60) = 1
controller.switch(61) = 1
controller.switch(62) = 1

Sub sw60_hit
sw60.Isdropped=True
Controller.switch (60)= 0
PlaySound "DTC"
End sub

Sub sw61_hit
sw61.Isdropped=True
Controller.switch (61)= 0
PlaySound "DTC"
End sub

Sub sw62_hit
sw62.Isdropped=True
Controller.switch (62)= 0
PlaySound "DTC"
End sub

Sub resetinline(Enabled)
	If Enabled Then
		sw60.Isdropped=false
		Controller.switch (60)=1
		sw61.Isdropped=False
		Controller.switch (61)=1
		sw62.Isdropped=False
		Controller.switch (62)=1
		PlaySound SoundFX("DTReset",DOFContactors)
	End If
 End Sub

Sub SnakeKick(enabled)
    If enabled Then
	sw54.timerenabled = true
    End If
End Sub

Sub SnakeJawLatch(enabled)
	If Enabled Then
	snakejawf.rotatetoend
	JawLatch.isDropped = 0
	Controller.Switch(56) = 0
	End If
End Sub

Sub Solsnakejaw(enabled)
	If Enabled Then
	snakejawf.rotatetoStart
	JawLatch.isDropped = 1
	Controller.Switch(56) = 1
	End If
End Sub

Sub sw54_Timer()
	PlaySound SoundFX("popper_ball", DOFContactors)
    sw54.Kick 195, 32
    controller.switch (54) = 0
	sw54.timerenabled = 0
End Sub

Sub Hammer(enabled)
	If Enabled Then
	HammerTD.Enabled = 1
	HammerTU.Enabled = 0
	MagnetHole.Enabled = 1
	HMag.MagnetOn = 0
	Else
	HammerTD.Enabled = 0
	HammerTU.Enabled = 1
	MagnetHole.Enabled = 0
	End If
End Sub

Sub HammerTD_Timer
	If HammerTU.Enabled = 0 Then
		If HammerP.RotX > 0 Then
			HammerP.RotX = HammerP.RotX -1
		ElseIf HammerP.RotX = 0 Then
			HammerTD.enabled = 0
		End If
	End If
End Sub

Sub HammerTU_Timer
	If HammerTD.Enabled = 0 Then
		If HammerP.RotX < 35 Then
			Hammerp.RotX = HammerP.RotX + 1
		ElseIf HammerP.RotX = 35 Then
			HammerTU.Enabled = 0
		End If
	End If
End Sub

Sub CoffinMagDown(enabled)
	if Enabled Then
	HMagnetP.Z = -75
	MagnetHole.Enabled = 1
	Else
	HMagnetP.Z = 0
	MagnetHole.Enabled = 0
	End If
End Sub

Sub CrossMotor(enabled)
	If Enabled Then
		If CrossPrim.Z = 60 Then CrossTimerDown.Enabled = 1
		If CrossPrim.Z = 0 Then CrossTimerUp.Enabled = 1
	Else
	End If
End Sub


Sub CrossTimerUp_Timer
		If CrossPrim.Z < 60 Then
			CrossPrim.Z = CrossPrim.Z+1
			F119.bulbhaloheight = f119.bulbhaloheight+1
			F119b.bulbhaloheight = f119b.bulbhaloheight+1
			F119c.bulbhaloheight = f119c.bulbhaloheight+1
			F119d.bulbhaloheight = f119d.bulbhaloheight+1
			Controller.Switch(33)=0
			Controller.Switch(34)=0
		 End If
		If CrossPrim.Z = 60 Then Controller.Switch(34) = 1:CrossTimerUp.Enabled = 0
End Sub

Sub CrossTimerDown_Timer
		If CrossPrim.Z > 0 Then
			CrossPrim.Z = CrossPrim.Z-1
			F119.bulbhaloheight = f119.bulbhaloheight-1
			F119b.bulbhaloheight = f119b.bulbhaloheight-1
			F119c.bulbhaloheight = f119c.bulbhaloheight-1
			F119d.bulbhaloheight = f119d.bulbhaloheight-1
			Controller.Switch(33)=0
			Controller.Switch(34)=0
		End If
		If CrossPrim.Z = 0 Then Controller.Switch(33) = 1: CrossTimerDown.Enabled = 0
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

Sub orbitpost(Enabled)
	If Enabled Then
		UpPost.Isdropped=0
	Else
		UpPost.Isdropped=1
	End If
 End Sub


Sub SolSparkyHead(Enabled)
    If Enabled Then
        SparkyShake
		Playsound "SmallSol"
    End If
End Sub

Sub SparkyShake
    cball.vely = 10 + 2 * (RND(1) - RND(1) )
End Sub

Sub Sparky
    Sparkyhead.rotx = (ckicker.y - cball.y)
    Sparkyhead.roty = (cball.x - ckicker.x)
End Sub

Sub SolLFlipper(Enabled)
     If Enabled Then
		 PlaySound SoundFX("FlipperUpLeft",DOFFlippers)
		 LeftFlipper.RotateToEnd
     Else
		 PlaySound SoundFX("FlipperDown",DOFFlippers)
		 LeftFlipper.RotateToStart
     End If
 End Sub

Sub SolRFlipper(Enabled)
     If Enabled Then
		 PlaySound SoundFX("FlipperUpRight",DOFFlippers)
		 RightFlipper.RotateToEnd
     Else
		 PlaySound SoundFX("FlipperDown",DOFFlippers)
		 RightFlipper.RotateToStart
     End If
 End Sub

 Dim dBall, dZpos

Sub ScoopKick(Enabled)
	If Enabled Then
	Kicker1.Enabled = 0
	vpmtimer.addtimer 600, "Kicker1.enabled= 1'"
	bsRHole.ExitSol_On
	End If
End Sub

 Sub kicker3_Hit
      bsRHole.AddBall Me
 End Sub

Sub Kicker1_Hit
	Playsound "ScoopLeft"
End Sub

Sub Flash19(Level)
	If Level > 0 Then
		F119.intensityScale = Level / 5
		f119b.intensityScale = Level / 5
		f119c.intensityScale = Level / 5
		f119d.intensityScale = Level / 5
		F119.State = 1
		F119b.State = 1
		F119c.State = 1
		F119d.State = 1
		CrossPrim.Image = "Crossmapon"
	Else
		F119.State = 0
		F119b.State = 0
		F119c.State = 0
		F119d.State = 0
		CrossPrim.Image = "Crosstex"
	end If
end Sub

Sub Flash23(Level)
	If Level > 0 Then
		F123.intensityScale = Level / 5
		F123b.intensityScale = Level / 5
		F123.State = 1
		F123b.State = 1
	Else
		F123.State = 0
		F123b.State = 0
	end If
end Sub

Sub Flash26(Level)
	If Level > 0 Then
		f126a.intensityScale = Level / 5
		F126b.intensityScale = Level / 5
		F126c.intensityScale = Level
		f126a.State = 1
		F126b.State = 1
		F126c.visible = 1
	Else
		f126a.State = 0
		F126b.State = 0
		F126c.visible = 0
	end If
end Sub

Sub Flash27(Level)
	If Level > 0 Then
		f127a.intensityScale = Level / 5
		F127b.intensityScale = Level / 5
		f127a.State = 1
		F127b.State = 1
	Else
		f127a.State = 0
		F127b.State = 0
	end If
end Sub

Sub Flash28(Level)
	For each xx in SparkySpots
	If Level > 0 Then
		xx.intensityScale = level /5
		xx.state = 1
	Else
		xx.state = 0
	End If
	Next
End Sub

Sub Flash29(Level)
	If Level > 0 Then
		f129.intensityScale = Level / 5
		F129b.intensityScale = Level / 5
		F129c.intensityScale = Level
		f129.State = 1
		F129b.State = 1
		F129c.visible = 1
	Else
		f129.State = 0
		F129b.State = 0
		F129c.visible = 0
	end If
end Sub

Sub Flash30(Level)
	If Level > 0 Then
		f130.intensityScale = Level / 5
		f130b.intensityScale = Level / 5
		f130c.intensityScale = Level
		f130d.intensityScale = Level
		f130.State = 1
		f130b.State = 1
		f130c.Visible = 1
		f130d.Visible = 1
	Else
		f130.State = 0
		f130b.State = 0
		f130c.Visible = 0
		f130d.Visible = 0
	end If
end Sub

Sub Flash31(Level)
	If Level > 0 Then
		If CoffinMod = 0 Then
			L31C.intensityScale = Level / 5
			L31C2.intensityScale = Level / 5
			l31f.visible = 0
			l31f2.Visible = 0
			L31C.Visible = 1
			L31C2.visible = 1
			L31C.State = 1
			L31C2.State = 1
		Elseif coffinMod = 2 then
			L31C.intensityScale = Level / 5
			L31C2.intensityScale = Level / 5
			l31f.State = 1
			l31f2.State = 1
			l31f.visible = 1
			l31f2.Visible = 1
			L31C.Visible =0
			L31C2.visible = 0
		End If
	Else
	l31f.State = 0
	l31f2.State = 0
	L31C.State = 0
	L31C2.State = 0
	l31f.visible = 0
	l31f2.visible = 0
	end If
end Sub

Sub Flash32(Level)
	If Level > 0 Then
		F132.intensityScale = Level / 5
		F132.State = 1
	Else
		F132.State = 0
	End If
End Sub

Sub CLockRelease(Enabled)
	If Enabled Then
	CoffinStop.isDropped = 1
	Else
	CoffinStop.isDropped = 0
	End If
End Sub

'***********************************************
'***********************************************
					'Switches
'***********************************************
'***********************************************

Sub CapWall_Hit:PlaySound "fx_collide":End Sub

Sub sw1_Hit:Controller.Switch(1) = 1:PlaySound "sensor":End Sub
Sub sw1_UnHit:Controller.Switch(1) = 0:End Sub
Sub sw3_Hit:Controller.Switch(3) = 1:PlaySound "sensor":End Sub
Sub sw3_UnHit:Controller.Switch(3) = 0:End Sub
Sub sw4_Hit:vpmTimer.PulseSw 4:PlaySound "target":End Sub
Sub sw7_Hit:vpmTimer.PulseSw 7:PlaySound "target":End Sub
Sub sw9_Hit:vpmTimer.PulseSw 9:PlaySound "target":End Sub
Sub sw23_Hit:Controller.Switch(23) = 1:PlaySound "sensor":End Sub
Sub sw23_UnHit:Controller.Switch(23) = 0:End Sub
Sub sw24_Hit:Controller.Switch(24) = 1:PlaySound "sensor"End Sub
Sub sw24_UnHit:Controller.Switch(24) = 0:End Sub
Sub sw25_Hit:Controller.Switch(25) = 1:PlaySound "sensor"End Sub
Sub sw25_UnHit:Controller.Switch(25) = 0:End Sub
Sub sw28_Hit:Controller.Switch(28) = 1:PlaySound "sensor":End Sub
Sub sw28_UnHit:Controller.Switch(28) = 0:End Sub
Sub sw29_Hit:Controller.Switch(29) = 1:PlaySound "sensor":End Sub
Sub sw29_UnHit:Controller.Switch(29) = 0:End Sub
Sub sw35_Hit:vpmTimer.PulseSw 35:PlaySound "target":End Sub
Sub sw36_Hit():playsound "target":vpmTimer.PulseSw 36:End Sub
Sub sw37_Hit():playsound "target":vpmTimer.PulseSw 37:End Sub
Sub Lspinner_Spin():vpmTimer.PulseSW 38:End Sub
Sub Rspinner_Spin():vpmTimer.PulseSW 39:End Sub
Sub sw40_Hit():playsound "target":vpmTimer.PulseSw 40:End Sub
Sub sw41_Hit():playsound "target":vpmTimer.PulseSw 41:End Sub

Sub sw42_Hit()
	playsound "target"
	vpmTimer.PulseSw 42
	sw42.HasHitEvent = 0
	sw42.timerenabled = 1
End Sub

sub sw42_timer
	sw42.HasHitEvent = 1
	sw42.timerenabled = 0
end Sub


Sub sw43_Hit:Controller.Switch(43) = 1:PlaySound "sensor":End Sub
Sub sw43_UnHit:Controller.Switch(43) = 0:End Sub
Sub sw44_Hit:Controller.Switch(44) = 1:PlaySound "sensor":End Sub
Sub sw44_UnHit:Controller.Switch(44) = 0:End Sub
Sub sw45_Hit:Controller.Switch(45) = 1:PlaySound "sensor":End Sub
Sub sw45_UnHit:Controller.Switch(45) = 0:End Sub
Sub sw46_Hit:Controller.Switch(46) = 1:PlaySound "sensor":End Sub
Sub sw46_UnHit:Controller.Switch(46) = 0:End Sub
Sub sw47g_hit:Playsound "Gate":vpmTimer.PulseSw 47:PlaySound "RightRampWire", 0, 1, Pan(ActiveBall):End Sub
Sub sw50g_hit:Playsound "Gate":vpmTimer.PulseSw 50:Playsound "LeftRampWire", 0, 1, Pan(ActiveBall):End Sub
Sub sw52_Hit:Controller.Switch(52)=1:End Sub
Sub sw52_unHit:Controller.Switch(52)=0:End Sub
Sub sw53_Hit:vpmTimer.PulseSw 53:End Sub
Sub sw54_Hit:Controller.Switch(54) = 1:Playsound "kicker_enter_center": End Sub
Sub JawLatch_Hit:playsound "target":Controller.Switch(55) = 1:me.TimerEnabled = 1:End Sub
Sub JawLatch_Timer:Controller.Switch(55) = 0: me.timerenabled = 0:End Sub
Sub sw57_Hit:Controller.Switch(57)=1:End Sub
Sub sw57_unHit:Controller.Switch(57)=0:End Sub
Sub sw58_Hit:Controller.Switch(58)=1:End Sub
Sub sw58_unHit:Controller.Switch(58)=0:End Sub
Sub sw59_Hit:Controller.Switch(59)=1:End Sub
Sub sw59_unHit:Controller.Switch(59)=0:End Sub
Sub SW63_Hit:Controller.Switch(63) = 1: End Sub
Sub sw63_unHit:Controller.Switch(63) = 0: End Sub
Sub LMagnet_Hit:LMag.AddBall ActiveBall:End Sub
Sub LMagnet_UnHit:LMag.RemoveBall ActiveBall:End Sub
Sub HMagnet_Hit:HMag.AddBall ActiveBall:End Sub
Sub HMagnet_UnHit:HMag.RemoveBall ActiveBall:End Sub
Sub SW64_Hit:Controller.Switch(64) = 1: sw64.timerenabled = 1:End Sub

Sub sw64_Timer
	If HMag.MagnetOn = 0 Then
	sw64.kick 100, 2
	Controller.Switch(64) = 0
	End If
	sw64.TimerEnabled = 0
End Sub


Sub RMagnet_Hit
	RMag.AddBall ActiveBall
	If RMag.MagnetOn=true then
	End If
End Sub

Sub RMagnet_UnHit:RMag.RemoveBall ActiveBall:End Sub

Dim BallCount:BallCount = 0
Sub Drain_Hit()
	PlaySound "Drain"
	BallCount = BallCount - 1
	bsTrough.AddBall Me
End Sub

Sub BallRelease_UnHit()
		BallCount = BallCount + 1
End Sub

Sub RLS_Timer()
	sw47p.RotX = -(sw47g.currentangle)
	sw50p.RotX = -(sw50g.currentangle)
	If Leftsling = True and Left1.ObjRotZ < -7 then Left1.ObjRotZ = Left1.ObjRotZ + 2
	If Leftsling = False and Left1.ObjRotZ > -20 then Left1.ObjRotZ = Left1.ObjRotZ - 2
	If Left1.ObjRotZ >= -7 then Leftsling = False
	If Leftsling = True and Left2.ObjRotZ > -212.5 then Left2.ObjRotZ = Left2.ObjRotZ - 2
	If Leftsling = False and Left2.ObjRotZ < -199 then Left2.ObjRotZ = Left2.ObjRotZ + 2
	If Left2.ObjRotZ <= -212.5 then Leftsling = False
	If Leftsling = True and Left3.TransZ > -23 then Left3.TransZ = Left3.TransZ - 4
	If Leftsling = False and Left3.TransZ < -0 then Left3.TransZ = Left3.TransZ + 4
	If Left3.TransZ <= -23 then Leftsling = False

	If Rightsling = True and Right1.ObjRotZ > 7 then Right1.ObjRotZ = Right1.ObjRotZ - 2
	If Rightsling = False and Right1.ObjRotZ < 20 then Right1.ObjRotZ = Right1.ObjRotZ + 2
	If Right1.ObjRotZ <= 7 then Rightsling = False
	If Rightsling = True and Right2.ObjRotZ < 212.5 then Right2.ObjRotZ = Right2.ObjRotZ + 2
	If Rightsling = False and Right2.ObjRotZ > 199 then Right2.ObjRotZ = Right2.ObjRotZ - 2
	If Right2.ObjRotZ >= 212.5 then Rightsling = False
	If Rightsling = True and Right3.TransZ > -23 then Right3.TransZ = Right3.TransZ - 4
	If Rightsling = False and Right3.TransZ < -0 then Right3.TransZ = Right3.TransZ + 4
	If Right3.TransZ <= -23 then Rightsling = False
End Sub

Sub LeftSlingShot_Slingshot
	Leftsling = True
	Controller.Switch(26) = 1
 	PlaySound Soundfx("left_slingshot",DOFContactors):LeftSlingshot.TimerEnabled = 1
  End Sub

Dim Leftsling:Leftsling = False


 Sub LeftSlingShot_Timer:Me.TimerEnabled = 0:Controller.Switch(26) = 0:End Sub

 Sub RightSlingShot_Slingshot
	Rightsling = True
	Controller.Switch(27) = 1
 	PlaySound Soundfx("right_slingshot",DOFContactors):RightSlingshot.TimerEnabled = 1
  End Sub

 Dim Rightsling:Rightsling = False

Sub RightSlingShot_Timer:Me.TimerEnabled = 0:Controller.Switch(27) = 0:End Sub

    Const IMPowerSetting = 55
    Const IMTime = 0.6
    Set plungerIM = New cvpmImpulseP
    With plungerIM
        .InitImpulseP swplunger, IMPowerSetting, IMTime
        .Random 0.3
        .InitExitSnd SoundFX("autoplunger" ,DOFContactors), SoundFX("plunger" ,DOFContactors)
        .CreateEvents "plungerIM"
    End With

      Sub Bumper1b_Hit
      vpmTimer.PulseSw 31
      PlaySound SoundFX("fx_bumper1",DOFContactors)
    	End Sub


      Sub Bumper2b_Hit
      vpmTimer.PulseSw 30
      PlaySound SoundFX("fx_bumper1",DOFContactors)
       End Sub

      Sub Bumper3b_Hit
      vpmTimer.PulseSw 32
      PlaySound SoundFX("fx_bumper1",DOFContactors)
       End Sub

 Dim LampState(300), FadingLevel(300), FadingState(300)
Dim FlashState(200), FlashLevel(200), FlashMin(200), FlashMax(200)
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

    FlashSpeedUp = 50   ' fast speed when turning on the flasher
    FlashSpeedDown = 10 ' slow speed when turning off the flasher, gives a smooth fading
    AllFlashOff()
End Sub

Sub AllFlashOff
    Dim i
    For i = 0 to 200
        FlashState(i) = 0
    Next
End Sub

Sub SetRGBLamp(Lamp, R, G, B)
	 Lamp.Color = RGB(R, G, B)
	 Lamp.ColorFull = RGB(R, G, B)
	 Lamp.State = 1
End Sub

 Sub UpdateLamps
  	NFadeL 17, l17
  	NFadeL 18, l18
  	NFadeL 19, l19
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
  	NFadeL 53, l53
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
  	NFadeL 67, l67
  	NFadeL 68, l68
  	NFadeL 69, l69
  	NFadeL 70, l70
  	NFadeL 71, l71
  	NFadeLm 73, l73
	NFadeL 73, l73f
  	NFadeLm 74, l74
  	NFadeL 74, l74f
  	NFadeLm 75, l75
  	NFadeL 75, l75f
  	NFadeL 77, l77
  	NFadeL 78, l78
  	NFadeL 79, l79
  	NFadeL 80, l80
	NFadeL 120, f120
	NFadeLm 121, l21a
	NFadeLm 121, l21b
	NFadeL 121, l21c
	NFadeLm 122, l22a
	NFadeLm 122, l22b
	NFadeL 122, l22c
	NFadeL 125, f125



SetRGBLamp CN4, LampState(87),LampState(88),LampState(89)
SetRGBLamp CN9, Lampstate(99),Lampstate(100),Lampstate(101)
SetRGBLamp CN5, Lampstate(90), Lampstate(91),Lampstate(92)
SetRGBLamp CN13, Lampstate(108),Lampstate(109),Lampstate(110)
SetRGBLamp CN11, Lampstate(102),Lampstate(103),Lampstate(104)
SetRGBLamp CN19, Lampstate(126),Lampstate(127),Lampstate(128)

	dim bulb, bulb2, bulb3, bulb4
	for each bulb in GIW
	SetRGBLamp bulb, LampState(136),LampState(136),LampState(136)
	next
	for each bulb2 in GIR
	SetRGBLamp bulb2, LampState(130),0,0
	next
	for each bulb3 in GIB
	SetRGBLamp bulb3, 0,0,LampState(132)
	next
	for each bulb4 in GIWUP
	SetRGBLamp bulb4, LampState(134),LampState(134),LampState(134)
	next

	If SideWallGI = 1 Then
		For each xx in GIWF: xx.intensityScale = LampState(136): Next
	End If

End Sub

Dim GILevel

Sub Intensity
	If DayNight <= 20 Then
			GILevel = .5
	ElseIf DayNight <= 40 Then
			GILevel = .4125
	ElseIf DayNight <= 60 Then
			GILevel = .325
	ElseIf DayNight <= 80 Then
			GILevel = .2375
	Elseif DayNight <= 100  Then
			GILevel = .15
	End If

	For each xx in GIW: xx.Intensity = xx.Intensity * GILevel: Next
	For each xx in GIR: xx.Intensity = xx.Intensity * GILevel: Next
	For each xx in GIB: xx.Intensity = xx.Intensity * GILevel: Next
	For each xx in GIWUP: xx.Intensity = xx.Intensity * GILevel: Next
End Sub

Sub FadePrim(nr, pri, a, b, c, d)
    Select Case FadingLevel(nr)
        Case 2:pri.image = d:FadingLevel(nr) = 0
        Case 3:pri.image = c:FadingLevel(nr) = 1
        Case 4:pri.image = b:FadingLevel(nr) = 2
        Case 5:pri.image = a:FadingLevel(nr) = 3
    End Select
End Sub

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

 End Sub

Set MotorCallback = GetRef("GameTimer")

Sub GameTimer
    Sparky
    RollingSound
End Sub


Sub Pins_Hit (idx)
	PlaySound "pinhit_low", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Targets_Hit (idx)
	PlaySound SoundFX("target",DOFContactors), 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub TargetBankWalls_Hit (idx)
	PlaySound "target", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Metals_Thin_Hit (idx)
	PlaySound "metalhit_thin", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
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

Sub Rubbers_Hit(idx)
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 20 then
		PlaySound "fx_rubber2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End if
	If finalspeed >= 6 AND finalspeed <= 20 then
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


'*****************************************
'			BALL SHADOW
'*****************************************
Dim BallShadow
BallShadow = Array (BallShadow1,BallShadow2,BallShadow3,BallShadow4,BallShadow5,BallShadow6,BallShadow7,BallShadow8)

Sub BallShadowUpdate_timer()
    Dim BOT, b
    BOT = GetBalls
    ' hide shadow of deleted balls
    If UBound(BOT)<(tnob-1) Then
        For b = (UBound(BOT) + 1) to (tnob-1)
            BallShadow(b).visible = 8
        Next
    End If
    ' exit the Sub if no balls on the table
    If UBound(BOT) = -1 Then Exit Sub
    ' render the shadow for each ball
    For b = 0 to UBound(BOT)
        If BOT(b).X < Table.Width/2 Then
            BallShadow(b).X = ((BOT(b).X) - (Ballsize/6) + ((BOT(b).X - (Table.Width/2))/7)) + 13
        Else
            BallShadow(b).X = ((BOT(b).X) + (Ballsize/6) + ((BOT(b).X - (Table.Width/2))/7)) - 13
        End If
        ballShadow(b).Y = BOT(b).Y + 10
        If BOT(b).Z > 20 Then
            BallShadow(b).visible = 1
        Else
            BallShadow(b).visible = 0
        End If
    Next
End Sub

'*****************************************
'			FLIPPER SHADOWS
'*****************************************

sub FlipperTimer_Timer()
	FlipperLSh.RotZ = LeftFlipper.currentangle
	FlipperRSh.RotZ = RightFlipper.currentangle

End Sub

Dim leftdrop:leftdrop = 0
Sub leftdrop1_Hit:leftdrop = 1:End Sub
Sub leftdrop2_Hit:
	If leftdrop = 1 then
		PlaySound "drop_left"
	End If
	StopSound "RightRamp"
	leftdrop = 0
End Sub

Dim rightdrop:rightdrop = 0
Sub rightdrop1_Hit:rightdrop = 1:End Sub
Sub rightdrop2_Hit
	If rightdrop = 1 then
		PlaySound "drop_Right"
	End If
	StopSound "RightRamp"
	rightdrop = 0
End Sub

Sub RRamp_Hit
	Playsound "RightRampMetal", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
End Sub

Sub LRamp_Hit
	Playsound "LeftRampMetal", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
End Sub

Sub PrimT_Timer
	snakejaw.objrotx = snakejawf.currentangle
	if l66.State = 1 then L66f.visible = 1 else l66f.visible = 0
	if l67.State = 1 then l67f.visible = 1 else l67f.visible = 0
	if l23.State = 1 then l23f.visible = 1 else l23f.visible = 0
	if l24.State = 1 then l24f.visible = 1 else l24f.visible = 0
	if l21a.State = 1 then f21a.visible = 1 else f21a.visible = 0
	if l21a.State = 1 then f21b.visible = 1 else f21b.visible = 0
	if l21a.State = 1 then f21c.visible = 1 else f21c.visible = 0
	if l22a.State = 1 then f22a.visible = 1 else f22a.visible = 0
	if l22a.State = 1 then f22b.visible = 1 else f22b.visible = 0
	if l22a.State = 1 then f22c.visible = 1 else f22c.visible = 0
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

Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "table" is the name of the table
  Dim tmp
  tmp = tableobj.y * 2 / table.height-1
  If tmp > 0 Then
    AudioFade = Csng(tmp ^10)
  Else
    AudioFade = Csng(-((- tmp) ^10) )
  End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "table" is the name of the table
  Dim tmp
  tmp = tableobj.x * 2 / table.width-1
  If tmp > 0 Then
    AudioPan = Csng(tmp ^10)
  Else
    AudioPan = Csng(-((- tmp) ^10) )
  End If
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

Function AudioFade(ball) ' Can this be together with the above function ?
  Dim tmp
  tmp = ball.y * 2 / table.height-1
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

Const tnob = 8 ' total number of balls
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
  If table.VersionMinor > 3 OR table.VersionMajor > 10 Then
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 200, Pan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
  Else
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 200, Pan(ball1), 0, Pitch(ball1), 0, 0
  End if
End Sub


' Thalamus : Exit in a clean and proper way
Sub table_exit()
  Controller.Pause = False
  Controller.Stop
End Sub


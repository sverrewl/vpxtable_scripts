'Paragon (Bally 1978) v2.0
' by bord
' base script by 32assassin

' Thalamus 2018-07-24
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Changed UseSolenoids=1 to 2
' Thalamus : 2018-08-16 : Improvied SSF positions.
' Remeber - first gate needs adjusting because of rubber changes to 10.5

' ! NOT Tested yet !

Option Explicit
Randomize

Const VolDiv = 4000

Const VolBump   = 1.5  ' Bumpers multiplier.
Const VolRol    = 1    ' Rollovers volume multiplier.
Const VolGates  = 1    ' Gates volume multiplier.
Const VolMetals = 1    ' Metals volume multiplier.
Const VolRB     = 1    ' Rubber bands multiplier.
Const VolRH     = 1    ' Rubber hits multiplier.
Const VolRPo    = 1    ' Rubber posts multiplier.
Const VolRPi    = 1    ' Rubber pins multiplier.
Const VolPlast  = 1    ' Plastics multiplier.
Const VolTarg   = 1    ' Targets multiplier.
Const VolWood   = 1    ' Woods multiplier.

Const VolSpin   = 1.5  ' Spinners volume.
Const VolSling  = 1.5  ' Spinners volume.
Const VolFlip   = 1    ' Flipper volume.


On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Const cGameName="paragon",UseSolenoids=2,UseLamps=1,UseGI=0,SSolenoidOn="SolOn",SSolenoidOff="SolOff", SCoin="coin"

LoadVPM "01130100", "Bally.VBS", 3.21

Dim BallShadows: Ballshadows=1  		'******************set to 1 to turn on Ball shadows
Dim FlipperShadows: FlipperShadows=1  '***********set to 1 to turn on Flipper shadows

'*********** Desktop/Cabinet settings ************************

Dim DesktopMode: DesktopMode = Table1.ShowDT
Dim hiddenvalue

If DesktopMode = True Then 'Show Desktop components
Ramp16.visible=1
Ramp15.visible=1
hiddenvalue=0
Else
Ramp16.visible=0
Ramp15.visible=0
hiddenvalue=1
End if

Sub SolLFlipper(Enabled)
     If Enabled Then
         PlaySoundAtVol SoundFX("flipperupleft",DOFContactors),LeftFlipper, VolFlip:LeftFlipper.RotateToEnd:LeftFlipper1.RotateToEnd
     Else
         PlaySoundAtVol SoundFX("fx_Flipperdown",DOFContactors),LeftFlipper, VolFlip:LeftFlipper.RotateToStart:LeftFlipper1.RotateToStart
     End If
  End Sub

Sub SolRFlipper(Enabled)
     If Enabled Then
         PlaySoundAtVol SoundFX("flipperupright",DOFContactors),RightFlipper,VolFlip:RightFlipper.RotateToEnd:RightFlipper1.RotateToEnd
     Else
         PlaySoundAtVol SoundFX("fx_Flipperdown",DOFContactors),RightFlipper,VolFlip:RightFlipper.RotateToStart:RightFlipper1.RotateToStart
     End If
End Sub
'**********************************************************************************************************

'Solenoid Controlled toys
'**********************************************************************************************************

'*****GI Lights On
dim xx
For each xx in GI:xx.State = 1: Next

'Sub RelayAC(enabled)
'	vpmNudge.SolGameOn enabled
'End Sub

	if ballshadows=1 then
		BallShadowUpdate.enabled=1
	  else
		BallShadowUpdate.enabled=0
	end if

	if flippershadows=1 then
		FlipperLSh.visible=1
		FlipperRSh.visible=1
	  else
		FlipperLSh.visible=0
		FlipperRSh.visible=0
	end if

'Primitive Flipper Code
Sub FlipperTimer_Timer
	LUFLogo.roty = LeftFlipper.currentangle  - 55
	LFPrim.rotz = LeftFlipper.currentangle
	RFPrim.rotz = RightFlipper.currentangle
    rfprim1.rotz = rightflipper.currentangle
	if FlipperShadows=1 then
		FlipperLsh.rotz= LeftFlipper.currentangle
		FlipperLsh1.rotz= LeftFlipper1.currentangle
		FlipperRsh.rotz= RightFlipper.currentangle
		FlipperRsh1.rotz= RightFlipper1.currentangle
	end if
End Sub

'*****************************************
'			BALL SHADOW
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
            BallShadow(b).X = ((BOT(b).X) - (Ballsize/16) + ((BOT(b).X - (Table1.Width/2))/17))' + 13
        Else
            BallShadow(b).X = ((BOT(b).X) + (Ballsize/16) + ((BOT(b).X - (Table1.Width/2))/17))' - 13
        End If
        ballShadow(b).Y = BOT(b).Y + 10
        If BOT(b).Z > 20 Then
            BallShadow(b).visible = 1
        Else
            BallShadow(b).visible = 0
        End If
    Next
End Sub

'**********************************************************************************************************

'Initiate Table
'**********************************************************************************************************
Dim bsTrough, bsSaucerL, bsSaucerM, bsSaucerR, dtLL, dtRR
Sub Table1_Init
	vpmInit me
	On Error Resume Next
	With Controller
		.GameName=cGameName
		If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
		.SplashInfoLine= "Paragon (Bally)"
		.HandleMechanics=0
		.HandleKeyboard=0
		.ShowDMDOnly=1
		.ShowFrame=0
		.ShowTitle=0
		.Hidden=hiddenvalue
		.Run
		If Err Then MsgBox Err.Description
	On Error Goto 0
	End With

    PinMAMETimer.Interval = PinMAMEInterval
    PinMAMETimer.Enabled = 1

'*************************************************************
'Solenoid Call backs
'**********************************************************************************************************

SolCallback(2) = "dtLL.SolDropUp"						    		' Single Drop Targets Reset
SolCallback(6) = "vpmSolSound SoundFX(""Knocker"",DOFKnocker),"		' Knocker
SolCallback(8) = "SolLeftOut"									' Treasure Chamber Saucer Release
SolCallback(9) = "SolCenterOut"									' Golden Cliffs Saucer Release
SolCallback(10) = "SolRightOut"								' Paragon Saucer Release
SolCallback(17) = "SolRightTargetReset"				        		' Right 3-bank Target Bank Reset
SolCallback(7) = "bsTrough.SolOut"						   			' Outhole Kicker (Ball Release)
SolCallBack(19) = "vpmNudge.SolGameOn"									' K1 Relay (Flipper Enable)

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

	vpmNudge.TiltSwitch  = swtilt
	vpmNudge.Sensitivity = 2
	vpmNudge.Tiltobj = Array(Bumper1,Bumper2,Bumper3,Bumper4,LeftSlingshot,RightSlingshot)

	' Trough
		Set bsTrough = New cvpmBallStack
		bsTrough.Initsw 0,8,0,0,0,0,0,0
		bsTrough.InitKick BallRelease, 0,2
		bsTrough.InitExitSnd SoundFX("ballrelease",DOFContactors), SoundFX("Solenoid",DOFContactors)
		bsTrough.Balls = 1

 '	 Saucer Mid Right
	Set bsSaucerR = New cvpmBallStack
		bsSaucerR.InitSaucer sw24,24,220,8
        bsSaucerR.KickAngleVar = 4
        bsSaucerR.KickForceVar = 1
		bsSaucerR.InitExitSnd SoundFX("kicker_release",DOFContactors), SoundFX("Solenoid",DOFContactors)
 		'bsSaucerR.CreateEvents "bsSaucerR", sw24

'	 Saucer Left
	Set bsSaucerL = New cvpmBallStack
		bsSaucerL.InitSaucer sw31,31,45,12
        bsSaucerL.KickAngleVar = 4
        bsSaucerL.KickForceVar = 1
		bsSaucerL.InitExitSnd SoundFX("kicker_release",DOFContactors), SoundFX("Solenoid",DOFContactors)
		'bsSaucerL.CreateEvents "bsSaucerL", sw31

 '	 Saucer Mid Left
	Set bsSaucerM = New cvpmBallStack
		bsSaucerM.InitSaucer sw32,32,135,8
        bsSaucerM.KickAngleVar = 4
        bsSaucerM.KickForceVar = 1
		bsSaucerM.InitExitSnd SoundFX("kicker_release",DOFContactors), SoundFX("Solenoid",DOFContactors)
 		'bsSaucerM.CreateEvents "bsSaucerM", sw32

 	Set dtLL=New cvpmDropTarget
		dtLL.InitDrop Array(sw1,sw2,sw3,sw4),Array(1,2,3,4)
		dtLL.InitSnd SoundFX("DTDrop",DOFContactors),SoundFX("DTReset",DOFContactors)

 	Set dtRR=New cvpmDropTarget
		dtRR.InitDrop Array(sw17,sw18,sw19),Array(17,18,19)
		dtRR.InitSnd SoundFX("DTDrop",DOFContactors),SoundFX("DTReset",DOFContactors)

dim xx

		For each xx in dtRightLights: xx.state=0:Next

 End Sub

'**********************************************************************************************************
'Plunger code
'**********************************************************************************************************

Sub Table1_KeyDown(ByVal KeyCode)
	If KeyDownHandler(keycode) Then Exit Sub
	If keycode = PlungerKey Then Plunger.Pullback: playsoundatvol "plungerpull", plunger, 1
End Sub

Sub Table1_KeyUp(ByVal KeyCode)
	If KeyUpHandler(keycode) Then Exit Sub
	If keycode = PlungerKey Then Plunger.Fire:PlaySoundatvol "plunger", plunger, 1
End Sub

'**********************************************************************************************************
 ' Drain hole
Sub Drain_Hit:playsoundatvol "drain", drain, 1:bsTrough.addball me:End Sub
Sub sw24_Hit:playsoundatvol "Kicker_Hit",sw24,1:bsSaucerR.AddBall 0:End Sub
Sub sw31_Hit:playsoundatvol "Kicker_Hit",sw31,1:bsSaucerL.AddBall 0:End Sub
Sub sw32_Hit:playsoundatvol "Kicker_Hit",sw32,1:bsSaucerM.AddBall 0:End Sub

'Drop Targets
 Sub Sw1_Dropped:dtLL.Hit 1 :End Sub
 Sub Sw2_Dropped:dtLL.Hit 2 :End Sub
 Sub Sw3_Dropped:dtLL.Hit 3 :End Sub
 Sub Sw4_Dropped:dtLL.Hit 4 :End Sub

 Sub Sw17_Dropped:dtRR.Hit 1 : GI_drop3.state=1 : End Sub
 Sub Sw18_Dropped:dtRR.Hit 2 : GI_drop2.state=1 : End Sub
 Sub Sw19_Dropped:dtRR.Hit 3 : GI_drop1.state=1 : End Sub

Sub SolRightTargetReset(enabled)
	dim xx
	if enabled then
		Playsound "TargetReset"
		dtRR.SolDropUp enabled
		For each xx in DTRightLights: xx.state=0:Next
	end if
End Sub

'Wire Triggers
Sub SW22_Hit:Controller.Switch(22)=1 : playsoundatvol "rollover", sw22, VolRol : End Sub
Sub SW22_unHit:Controller.Switch(22)=0:End Sub
Sub SW23_Hit:Controller.Switch(23)=1 : playsoundatvol "rollover", sw23, VolRol : End Sub
Sub SW23_unHit:Controller.Switch(23)=0:End Sub

'Star Triggers
Sub SW26_Hit:Controller.Switch(26)=1 : playsoundatvol "rollover", sw26, VolRol : End Sub
Sub SW26_unHit:Controller.Switch(26)=0:End Sub
Sub SW28_Hit:Controller.Switch(28)=1 : playsoundatvol "rollover", sw28, VolRol : End Sub
Sub SW28_unHit:Controller.Switch(28)=0:End Sub
Sub SW34_Hit:Controller.Switch(34)=1 : playsoundatvol "rollover", sw34, VolRol : End Sub
Sub SW34_unHit:Controller.Switch(34)=0:End Sub
Sub SW34a_Hit:Controller.Switch(34)=1 : playsoundatvol "rollover", sw34a, VolRol : End Sub
Sub SW34a_unHit:Controller.Switch(34)=0:End Sub

'Stand Up Targets
Sub sw29_Hit : vpmTimer.PulseSw(29) : playsoundatvol "target", sw29, 1 : End Sub
Sub sw30_Hit : vpmTimer.PulseSw(30) : playsoundatvol "target", sw30, 1 : End Sub

'Scoring Rubber
Sub sw27a_Hit : vpmTimer.PulseSw(27) : playsoundatvol "rubber_hit_3", LeafSwitch5, VolRB:End Sub
Sub sw27b_Hit : vpmTimer.PulseSw(27) : playsoundatvol "rubber_hit_3", LeafSwitch6, VolRB : End Sub
Sub sw34b_Hit : vpmTimer.PulseSw(34) : playsoundatvol "rubber_hit_3", sw18, VolRB : End Sub

'Spinners
Sub sw33_Spin:vpmTimer.PulseSw 33 : playsoundatvol "fx_spinner", sw33, VolSpin : End Sub

'Bumpers
Sub Bumper1_Hit : vpmTimer.PulseSw(37) : playsoundatvol SoundFX("rightbumper_hit",DOFContactors),bumper1,VolBump: End Sub
Sub Bumper2_Hit : vpmTimer.PulseSw(38) : playsoundatvol SoundFX("leftbumper_hit",DOFContactors),bumper2,VolBump: End Sub
Sub Bumper3_Hit : vpmTimer.PulseSw(39) : playsoundatvol SoundFX("leftbumper_hit",DOFContactors),bumper3,VolBump: End Sub
Sub Bumper4_Hit : vpmTimer.PulseSw(40) : playsoundatvol SoundFX("rightbumper_hit",DOFContactors),bumper4,VolBump: End Sub


'*******************************************************************************************************************
'*******************************************************************************************************************
'Map Lights to an array
'*******************************************************************************************************************
'******************************************************************************************************************

set Lights(1)=L1
set Lights(2)=L2
set Lights(3)=L3
set Lights(4)=L4
set Lights(5)=L5

set Lights(6)=L6
set Lights(22)=L22
set Lights(38)=L38
set Lights(43)=L43

set Lights(7)=L7
set Lights(23)=L23
set Lights(39)=L39
set Lights(55)=L55
set Lights(8)=L8
set Lights(24)=L24
set Lights(40)=L40

set Lights(9)=L9
set Lights(25)=L25
set Lights(41)=L41
set Lights(57)=L57
set Lights(10)=L10
set Lights(26)=L26
set Lights(42)=L42

set Lights(12)=L12
set Lights(13)=L13
set Lights(14)=L14
set Lights(15)=L15
set Lights(17)=L17
set Lights(18)=L18
set Lights(19)=L19
set Lights(20)=L20
set Lights(21)=L21

set Lights(28)=L28
set Lights(30)=L30
set Lights(31)=L31
set Lights(33)=L33
set Lights(34)=L34
set Lights(35)=L35
set Lights(36)=L36
set Lights(37)=L37

set Lights(44)=L44
set Lights(46)=L46
set Lights(47)=L47
set Lights(49)=L49
set Lights(50)=L50
set Lights(51)=L51

set Lights(52)=L52
set Lights(53)=L53
set Lights(54)=L54

set Lights(56)=L56
set Lights(58)=L58
set Lights(59)=L59
set Lights(60)=L60
set Lights(62)=L62
set Lights(63)=L63

'*******************************************************************************************************************
'*******************************************************************************************************************
'*******************************************************************************************************************

Sub editDips
	Dim vpmDips : Set vpmDips = New cvpmDips
	With vpmDips
		.AddForm 700,400,"Paragon - DIP switches"
		.AddChk 2,10,180,Array("Match feature",&H00100000)'dip 21
		.AddChk 205,10,115,Array("Credits display",&H00080000)'dip 20
		.AddFrame 2,30,190,"Maximum credits",&H00070000,Array("5 credits",0,"10 credits",&H00010000,"15 credits",&H00020000,"20 credits",&H00030000,"25 credits",&H00040000,"30 credits",&H00050000,"35 credits",&H00060000,"40 credits",&H00070000)'dip 17&18&19
		.AddFrame 2,160,190,"Sound features",&H80000080,Array("chime effects",&H80000000,"chime and tunes",0,"noise",&H00000080,"noises and tunes",&H80000080)'dip 8&32
		.AddFrame 2,235,190,"High score to date",&H00000060,Array("no award",0,"1 credit",&H00000020,"2 credits",&H00000040,"3 credits",&H00000060)'dip 6&7
		.AddFrame 2,310,190,"High score feature",&H00006000,Array("no award",0,"extra ball",&H00004000,"replay",&H00006000)'dip 14&15
		.AddFrame 205,30,190,"Balls per game", 32768,Array("3 balls",0,"5 balls", 32768)'dip 16
		.AddFrame 205,76,190,"Hitting in line targets",&H00200000,Array("does not spot any letters",0,"spots any PARAGON letter",&H00200000)'dip 22
		.AddFrame 205,122,190,"PARAGON saucer adjustment",&H00400000,Array ("500 and 1 bonus advance",0,"3,000 and 2 bonus advances",&H00400000)'dip 23
		.AddFrame 205,168,190,"PARAGON saucer lites",&H00800000,Array("keep scanning",0,"do not scan",&H00800000)'dip 24
		.AddFrame 205,214,190,"3 drop target special can be collected",&H40000000,Array("more than 1 time",0,"only one time, then 25,000 lites",&H40000000)'dip 31
		.AddFrame 205,260,190,"Waterfall lane 5,000 lite",&H10000000,Array("is off at start of game",0,"is on at start of game",&H10000000)'dip 29
		.AddFrame 205,306,190,"In line drop targets",&H20000000,Array("1,000 and 1 bonus advance",0,"3,000 and 2 bonus advances",&H20000000)'dip 30
		.AddLabel 50,390,300,20,"After hitting OK, press F3 to reset game with new settings."
		.ViewDips
	End With
End Sub
Set vpmShowDips = GetRef("editDips")


'*******************************************************************************************************************
'*******************************************************************************************************************
'*******************************************************************************************************************

'Start of VPX Call Backs

'*******************************************************************************************************************
'*******************************************************************************************************************
'*******************************************************************************************************************


'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep, V1step, V2step, V3step, LKick, CKick, RKick

Sub RightSlingShot_Slingshot
	vpmtimer.pulsesw 35
    PlaySoundatvol SoundFX("rightslingshot",DOFContactors), sling1, VolSling
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.TransZ = -34
    RStep = 0
    RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.TransZ = -10
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.TransZ = 0:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
	vpmtimer.pulsesw 36
    PlaySoundatvol SoundFX("leftslingshot",DOFContactors), sling2, VolSling
    LSling.Visible = 0
    LSling1.Visible = 1
    sling2.TransZ = -32
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.TransZ = -10
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.TransZ = 0:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

Sub Vhit1_hit
    RubberLA.Visible = 0
    RubberLA1.Visible = 1
    V1Step = 0
    Vhit1.TimerEnabled = 1
End Sub

Sub Vhit1_Timer
    Select Case V1Step
        Case 3:RubberLA1.Visible = 0:RubberLA2.Visible = 1
        Case 4:RubberLA2.Visible = 0:RubberLA.Visible = 1:Vhit1.TimerEnabled = 0
    End Select
    V1Step = V1Step + 1
End Sub

Sub Vhit2_hit
    RubberLB.Visible = 0
    RubberLB1.Visible = 1
    V2Step = 0
    Vhit2.TimerEnabled = 1
End Sub

Sub Vhit2_Timer
    Select Case V2Step
        Case 3:RubberLB1.Visible = 0:RubberLB2.Visible = 1
        Case 4:RubberLB2.Visible = 0:RubberLB.Visible = 1:Vhit2.TimerEnabled = 0
    End Select
    V2Step = V2Step + 1
End Sub

Sub vhit3_hit
    lsling.Visible = 0
    lsling3.Visible = 1
    V3Step = 0
    vhit3.TimerEnabled = 1
End Sub

Sub vhit3_Timer
    Select Case V3Step
        Case 3:lsling3.Visible = 0:lsling4.Visible = 1
        Case 4:lsling4.Visible = 0:lsling.Visible = 1:vhit3.TimerEnabled = 0
    End Select
    V3Step = V3Step + 1
End Sub

Sub SolLeftOut(enabled)
	If enabled Then
		bsSaucerL.ExitSol_On
		LKick = 0
		kickarmtop_prim2.ObjRotX = -12
		LKickTimer.Enabled = 1
	End If
End Sub

Sub LKickTimer_Timer
    Select Case LKick
        Case 1:kickarmtop_prim2.ObjRotX = -50
        Case 2:kickarmtop_prim2.ObjRotX = -50
        Case 3:kickarmtop_prim2.ObjRotX = -50
        Case 4:kickarmtop_prim2.ObjRotX = -50
        Case 5:kickarmtop_prim2.ObjRotX = -50
        Case 6:kickarmtop_prim2.ObjRotX = -50
        Case 7:kickarmtop_prim2.ObjRotX = -50
        Case 8:kickarmtop_prim2.ObjRotX = -50
        Case 9:kickarmtop_prim2.ObjRotX = -50
        Case 10:kickarmtop_prim2.ObjRotX = -50
        Case 11:kickarmtop_prim2.ObjRotX = -24
        Case 12:kickarmtop_prim2.ObjRotX = -12
        Case 13:kickarmtop_prim2.ObjRotX = 0:LKickTimer.Enabled = 0
    End Select
    LKick = LKick + 1
End Sub

Sub SolCenterOut(enabled)
	If enabled Then
		bsSaucerM.ExitSol_On
		CKick = 0
		kickarmtop_prim1.ObjRotX = -12
		CKickTimer.Enabled = 1
	End If
End Sub

Sub CKickTimer_Timer
    Select Case CKick
        Case 1:kickarmtop_prim1.ObjRotX = -50
        Case 2:kickarmtop_prim1.ObjRotX = -50
        Case 3:kickarmtop_prim1.ObjRotX = -50
        Case 4:kickarmtop_prim1.ObjRotX = -50
        Case 5:kickarmtop_prim1.ObjRotX = -50
        Case 6:kickarmtop_prim1.ObjRotX = -50
        Case 7:kickarmtop_prim1.ObjRotX = -50
        Case 8:kickarmtop_prim1.ObjRotX = -50
        Case 9:kickarmtop_prim1.ObjRotX = -50
        Case 10:kickarmtop_prim1.ObjRotX = -50
        Case 11:kickarmtop_prim1.ObjRotX = -24
        Case 12:kickarmtop_prim1.ObjRotX = -12
        Case 13:kickarmtop_prim1.ObjRotX = 0:CKickTimer.Enabled = 0
    End Select
    CKick = CKick + 1
End Sub

Sub SolRightOut(enabled)
	If enabled Then
		bsSaucerR.ExitSol_On
		RKick = 0
		kickarmtop_prim.ObjRotX = -12
		RKickTimer.Enabled = 1
	End If
End Sub

Sub RKickTimer_Timer
    Select Case RKick
        Case 1:kickarmtop_prim.ObjRotX = -50
        Case 2:kickarmtop_prim.ObjRotX = -50
        Case 3:kickarmtop_prim.ObjRotX = -50
        Case 4:kickarmtop_prim.ObjRotX = -50
        Case 5:kickarmtop_prim.ObjRotX = -50
        Case 6:kickarmtop_prim.ObjRotX = -50
        Case 7:kickarmtop_prim.ObjRotX = -50
        Case 8:kickarmtop_prim.ObjRotX = -50
        Case 9:kickarmtop_prim.ObjRotX = -50
        Case 10:kickarmtop_prim.ObjRotX = -50
        Case 11:kickarmtop_prim.ObjRotX = -24
        Case 12:kickarmtop_prim.ObjRotX = -12
        Case 13:kickarmtop_prim.ObjRotX = 0:RKickTimer.Enabled = 0
    End Select
    RKick = RKick + 1
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


' Sub Pins_Hit (idx)
' 	PlaySound "pinhit_low", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
' End Sub
'
' Sub Targets_Hit (idx)
' 	PlaySound "target", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
' End Sub
'
' Sub Metals_Thin_Hit (idx)
' 	PlaySound "metalhit_thin", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
' End Sub
'
' Sub Metals_Medium_Hit (idx)
' 	PlaySound "metalhit_medium", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
' End Sub
'
' Sub Metals2_Hit (idx)
' 	PlaySound "metalhit2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
' End Sub
'
' Sub Gates_Hit (idx)
' 	PlaySound "gate4", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
' End Sub
'
' Sub Spinner_Spin
' 	PlaySound "fx_spinner",0,.25,0,0.25
' End Sub
'
' Sub Rubbers_Hit(idx)
'  	dim finalspeed
'   	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
'  	If finalspeed > 12 then
' 		PlaySound "fx_rubber2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
' 	End if
' 	If finalspeed >= 3 AND finalspeed <= 12 then
'  		RandomSoundRubber()
'  	End If
' End Sub
'
' Sub Posts_Hit(idx)
'  	dim finalspeed
'   	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
'  	If finalspeed > 12 then
' 		PlaySound "fx_rubber2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
' 	End if
' 	If finalspeed >= 3 AND finalspeed <= 12 then
'  		RandomSoundRubber()
'  	End If
' End Sub
'
' Sub RandomSoundRubber()
' 	Select Case Int(Rnd*3)+1
' 		Case 1 : PlaySound "rubber_hit_1", 2
' 		Case 2 : PlaySound "rubber_hit_2", 2
' 		Case 3 : PlaySound "rubber_hit_3", 2
' 	End Select
' End Sub
'
' Sub aRollovers_Hit(idx):PlaySound "fx_sensor", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
'
' Sub aPlastics_Hit(idx):PlaySound "fx_PlasticHit", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
'
' Sub aWoods_Hit(idx):PlaySound "fx_Woodhit", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub

Sub Pins_Hit (idx)
	PlaySound "pinhit_low", 0, Vol(ActiveBall)*VolRPi, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Targets_Hit (idx)
	PlaySound "target", 0, Vol(ActiveBall)*VolTarg, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Metals_Thin_Hit (idx)
	PlaySound "metalhit_thin", 0, Vol(ActiveBall)*VolMetals, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Metals_Medium_Hit (idx)
	PlaySound "metalhit_medium", 0, Vol(ActiveBall)*VolMetals, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Metals2_Hit (idx)
	PlaySound "metalhit2", 0, Vol(ActiveBall)*VolMetals, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Gates_Hit (idx)
	PlaySound "gate4", 0, Vol(ActiveBall)*VolGates, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Spinner_Spin
	PlaySound "fx_spinner",0,.25,0,0.25
End Sub

Sub Rubbers_Hit(idx)
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 12 then
		PlaySound "fx_rubber2", 0, Vol(ActiveBall)*VolRB, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End if
	If finalspeed >= 3 AND finalspeed <= 12 then
 		RandomSoundRubber()
 	End If
End Sub

Sub Posts_Hit(idx)
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 12 then
		PlaySound "fx_rubber2", 0, Vol(ActiveBall)*VolRPo, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End if
	If finalspeed >= 3 AND finalspeed <= 12 then
 		RandomSoundRubber()
 	End If
End Sub

Sub RandomSoundRubber()
	Select Case Int(Rnd*3)+1
		Case 1 : PlaySound "rubber_hit_1", VolRB
		Case 2 : PlaySound "rubber_hit_2", VolRB
		Case 3 : PlaySound "rubber_hit_3", VolRB
	End Select
End Sub

Sub aRollovers_Hit(idx):PlaySound "fx_sensor", 0, Vol(ActiveBall)*VolRol, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub

Sub aPlastics_Hit(idx):PlaySound "fx_PlasticHit", 0, Vol(ActiveBall)*VolPlast, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub

Sub aWoods_Hit(idx):PlaySound "fx_Woodhit", 0, Vol(ActiveBall)*VolWood, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub


Sub LeftFlipper_Collide(parm)
 	RandomSoundFlipper()
End Sub

Sub RightFlipper_Collide(parm)
 	RandomSoundFlipper()
End Sub

Sub RandomSoundFlipper()
	Select Case Int(Rnd*3)+1
		Case 1 : PlaySound "flip_hit_1", 2
		Case 2 : PlaySound "flip_hit_2", 2
		Case 3 : PlaySound "flip_hit_3", 2
	End Select
End Sub

Sub table1_Paused:Controller.Pause = 1:End Sub
Sub table1_unPaused:Controller.Pause = 0:End Sub
Sub table1_Exit:Controller.Stop:End Sub

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


' Thalamus : Exit in a clean and proper way
Sub Table1_exit()
  Controller.Pause = False
  Controller.Stop
End Sub


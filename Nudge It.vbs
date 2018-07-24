
option explicit
Randomize

' Thalamus 2018-07-24
' Table doesn't have standard "Positional Sound Playback Functions" or "Supporting Ball & Sound Functions"
' No special SSF tweaks yet.

Const BallSize = 90
Const swStartButton = 3
Const cGameName="nudgeit"
Const UseSolenoids=1,UseLamps=True,UseGI=0,UseSyn=1,SSolenoidOn="SolOn",SSolenoidOff="Soloff"
Const SCoin="coin3",cCredits="Nudge It"

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "Can't open controller.vbs"
On Error Goto 0
LoadVPM "01320000","gts3.vbs",3.1

Dim VarHidden
If Table1.ShowDT = true then
    VarHidden = 0
else
    VarHidden = 1
end if
'if B2SOn = true then VarHidden = 1

'***************************************************************
'*   				   Keyboard Handlers        	    	   *
'***************************************************************

Sub Table1_KeyDown(ByVal KeyCode)
	If KeyCode=LeftFlipperKey Then Controller.Switch(4)=1
	If KeyCode=RightFlipperKey Then Controller.Switch(5)=1
If KeyCode=PlungerKey Then Plunger.PullBack:Controller.Switch(10)=True
If vpmKeyDown(KeyCode) Then Exit Sub
End Sub

Sub Table1_KeyUp(ByVal KeyCode)
	If KeyCode=LeftFlipperKey Then Controller.Switch(4)=0
	If KeyCode=RightFlipperKey Then Controller.Switch(5)=0
If KeyCode=PlungerKey Then Plunger.Fire:PlaySound"Plunger":Controller.Switch(10)=False
If vpmKeyUp(KeyCode) Then Exit Sub
End Sub

'***************************************************************
'*   			       	   Solenoids        	    	       *
'***************************************************************

'SolCallback(17)="flasher17"
'SolCallback(18)="vpmFlasher Light672,"
'SolCallback(19)="vpmFlasher Light673,"
'SolCallback(20)="vpmFlasher Light674,"
'SolCallback(21)="vpmFlasher Light675,"
'SolCallback(22)="vpmFlasher Light676,"
'SolCallback(23)="vpmSolSound ""bell"","
'SolCallback(24)=""									'coin meter
'SolCallback(26)="vpmFlasher BGFlasher,"
'SolCallback(27)=""									'ticket dispenser
SolCallback(28)="solpin"
'SolCallback(30)="vpmSolSound ""knocker"","
'SolCallback(31)="vpmNudge.SolGameOn"				'Tilt Relay (T)


'***************************************************************
'*   				        Ball Gate           	    	   *
'***************************************************************
dim foon, fo
Sub Solpin(Enabled)
     If Enabled Then
         pin.transy = -45
         pin.Collidable = 0
     Else
         pin.transy = 0
         pin.Collidable = 1
     End If
End Sub

'***************************************************************
'*   				    Animate Top Gate                	   *
'***************************************************************

Sub GateTimer_Timer()
   Gate1P.RotZ = ABS(Gate1.currentangle)
End Sub	

'***************************************************************
'*   				      Table Init        	        	   *
'***************************************************************

Sub Table1_Init
	On Error Resume Next
	With Controller
		.GameName=cGameName
		If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description:Exit Sub
		'.SplashInfoLine="Nudge It" & vbnewline & "Table by Destruk"
		.ShowFrame=0
		.ShowTitle=0
		.ShowDMDOnly=1
		.HandleMechanics=0
		.HandleKeyboard=0
        .Hidden = VarHidden
		.Games(cGameName).Settings.Value("dmd_pos_x")=128
		.Games(cGameName).Settings.Value("dmd_pos_y")=330
		.Games(cGameName).Settings.Value("dmd_width")=300
		.Games(cGameName).Settings.Value("dmd_height")=65

	End With
	Controller.Run
	If Err Then MsgBox Err.Description
	On Error Goto 0

	PinMAMETimer.Interval=PinMAMEInterval
	PinMAMETimer.Enabled=1
	
	'vpmNudge.TiltSwitch=151
	'vpmNudge.TiltObj=Array()
	'vpmNudge.Sensitivity=3
		
	vpmMapLights ALights
	Kicker.CreateSizedBall(29)
	Kicker.Kick 90,1

End Sub



'***************************************************************
'*   				   Rollover Switches        	    	   *
'***************************************************************

Sub sw1_Hit():PlaySound "rollover", 0:Controller.Switch(11) = 1:End Sub
Sub sw1_UnHit():Controller.Switch(11) = 0:End Sub
Sub sw2_Hit():PlaySound "rollover", 0:Controller.Switch(12) = 1:End Sub
Sub sw2_UnHit():Controller.Switch(12) = 0:End Sub
Sub sw3_Hit():PlaySound "rollover", 0:Controller.Switch(13) = 1:End Sub
Sub sw3_UnHit():Controller.Switch(13) = 0:End Sub
Sub sw4_Hit():PlaySound "rollover", 0:Controller.Switch(14) = 1:End Sub
Sub sw4_UnHit():Controller.Switch(14) = 0:End Sub
Sub sw5_Hit():PlaySound "rollover", 0:Controller.Switch(15) = 1:End Sub
Sub sw5_UnHit():Controller.Switch(15) = 0:End Sub
Sub sw6_Hit():PlaySound "rollover", 0:Controller.Switch(16) = 1:End Sub
Sub sw6_UnHit():Controller.Switch(16) = 0:End Sub
Sub sw7_Hit:Controller.Switch(7) = 1:End Sub
Sub sw17_Hit():PlaySound "rollover", 0:Controller.Switch(17) = 1:End Sub
Sub sw17_UnHit():Controller.Switch(17) = 0:End Sub






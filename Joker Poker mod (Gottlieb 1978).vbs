Option Explicit
Randomize

Const cGameName="jokrpokr",UseSolenoids=2,UseLamps=1,UseGI=0,SSolenoidOn="SolOn",SSolenoidOff="SolOff",SFlipperOn="fx_Flipperup",SFlipperOff="fx_Flipperdown"
Const SCoin="coin"

' Thalamus 2018-07-23
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
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
Const VolGates  = 1    ' Gates volume.
Const VolMetal  = 1    ' Metals volume.
Const VolRH     = 1    ' Rubber hits volume.
Const VolPo     = 1    ' Rubber posts volume.
Const VolPi     = 1    ' Rubber pins volume.
Const VolTarg   = 1    ' Targets volume.
Const VolSpin   = 1.5  ' Spinners volume.
Const VolFlip   = 1    ' Flipper volume.


LoadVPM"01150000","GTS1.VBS",3.10
Dim DesktopMode: DesktopMode = Table1.ShowDT

'******************* Options *********************
' DMD/Backglass Controller Setting
Const cController = 3		'0=Use value defined in cController.txt, 1=VPinMAME, 2=UVP server, 3=B2S server
'*************************************************

'*************************************************************
Dim cNewController
Sub LoadVPM(VPMver, VBSfile, VBSver)
	Dim FileObj, ControllerFile, TextStr

	On Error Resume Next
	If ScriptEngineMajorVersion < 5 Then MsgBox "VB Script Engine 5.0 or higher required"
	ExecuteGlobal GetTextFile(VBSfile)
	If Err Then MsgBox "Unable to open " & VBSfile & ". Ensure that it is in the same folder as this table. " & vbNewLine & Err.Description

	cNewController = 1
	If cController = 0 then
		Set FileObj=CreateObject("Scripting.FileSystemObject")
		If Not FileObj.FolderExists(UserDirectory) then
			Msgbox "Visual Pinball\User directory does not exist. Defaulting to vPinMame"
		ElseIf Not FileObj.FileExists(UserDirectory & "cController.txt") then
			Set ControllerFile=FileObj.CreateTextFile(UserDirectory & "cController.txt",True)
			ControllerFile.WriteLine 1: ControllerFile.Close
		Else
			Set ControllerFile=FileObj.GetFile(UserDirectory & "cController.txt")
			Set TextStr=ControllerFile.OpenAsTextStream(1,0)
			If (TextStr.AtEndOfStream=True) then
				Set ControllerFile=FileObj.CreateTextFile(UserDirectory & "cController.txt",True)
				ControllerFile.WriteLine 1: ControllerFile.Close
			Else
				cNewController=Textstr.ReadLine: TextStr.Close
			End If
		End If
	Else
		cNewController = cController
	End If

	Select Case cNewController
		Case 1
			Set Controller = CreateObject("VPinMAME.Controller")
			If Err Then MsgBox "Can't Load VPinMAME." & vbNewLine & Err.Description
			If VPMver>"" Then If Controller.Version < VPMver Or Err Then MsgBox "VPinMAME ver " & VPMver & " required."
			If VPinMAMEDriverVer < VBSver Or Err Then MsgBox VBSFile & " ver " & VBSver & " or higher required."
		Case 2
			Set Controller = CreateObject("UltraVP.BackglassServ")
		Case 3,4
			Set Controller = CreateObject("B2S.Server")
	End Select
	On Error Goto 0
End Sub
'*************************************************************

'Solenoid Call backs
'**********************************************************************************************************
SolCallback(1)="bsTrough.SolOut"
SolCallback(2)="vpmSolSound ""Knocker"","
SolCallback(3)="vpmSolSound ""10"","
SolCallback(4)="vpmSolSound ""100"","
SolCallback(5)="vpmSolSound ""1000"","
SolCallback(6)="dtjj.SolDropUp"
SolCallback(7)="dtqq.SolDropUp"
SolCallback(8)="dtaa.SolDropUp"
SolCallback(17)="vpmNudge.SolGameOn"
SolCallback(sLLFlipper)="vpmSolFlipper LeftFlipper,nothing,"
SolCallback(sLRFlipper)="vpmSolFlipper RightFlipper,RightFlipper1,"

'**********************************************************************************************************

'Solenoid Controlled toys
'**********************************************************************************************************
'Primitive Flipper Code
Sub FlipperTimer_Timer
	Lflip.roty = LeftFlipper.currentangle
	RFlip.roty = RightFlipper.currentangle
	Rflip1.roty = RightFlipper1.currentangle -90
End Sub
'**********************************************************************************************************

'Initiate Table
'**********************************************************************************************************
Dim bsTrough, dtjj, dtqq, dtaa, dtkk

Sub Table1_Init
	vpmInit Me
	On Error Resume Next
		With Controller
		.GameName = cGameName
		If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
		.SplashInfoLine = "Joker Poker (Gottlieb)"&chr(13)&"You Suck"
		.HandleMechanics=0
		.HandleKeyboard=0
		.ShowDMDOnly=1
		.ShowFrame=0
		.ShowTitle=0
        .hidden = 1
		If Err Then MsgBox Err.Description
	End With
	On Error Goto 0
		Controller.SolMask(0)=0
      vpmTimer.AddTimer 2000,"Controller.SolMask(0)=&Hffffffff'" 'ignore all solenoids - then add the timer to renable all the solenoids after 2 seconds
		Controller.Run
	If Err Then MsgBox Err.Description
	On Error Goto 0

	PinMAMETimer.Interval=PinMAMEInterval
	PinMAMETimer.Enabled=1

     vpmNudge.TiltSwitch = 4
     vpmNudge.Sensitivity = 2
     vpmNudge.TiltObj = Array(Bumper1, Bumper2, LeftSlingshot)

  ' Trough
    Set bsTrough = New cvpmBallStack
		bsTrough.initnotrough ballrelease, 66, 90, 10
		bsTrough.InitExitSnd "ballrelease","solenoid"

  ' Drop targets
     set dtjj=New cvpmDropTarget
		 dtjj.InitDrop Array(sw21,sw22,sw23),Array(21,22,23)
         dtjj.initsnd "DTDrop","DTReset"

     set dtqq=New cvpmDropTarget
         dtqq.InitDrop Array(sw31,sw32,sw33),array(31,32,33)
         dtqq.initsnd "DTDrop","DTReset"

     set dtkk=New cvpmDropTarget
         dtkk.InitDrop Array(sw41,sw42,sw43,sw44),array(41,42,43,44)
         dtkk.initsnd "DTDrop","DTReset"

     set dtaa=New cvpmDropTarget
         dtaa.InitDrop Array(sw50,sw51,sw52,sw53,sw54),array(50,51,52,53,54)
         dtaa.initsnd "DTDrop","DTReset"

  End Sub

'**********************************************************************************************************
'Plunger code
'**********************************************************************************************************

Sub Table1_KeyDown(ByVal KeyCode)
	If KeyDownHandler(keycode) Then Exit Sub
	If keycode = PlungerKey Then Plunger.Pullback:playsoundAtVol"plungerpull", Plunger, 1
	If keycode=AddCreditKey then vpmTimer.pulseSW (swCoin1)
End Sub

Sub Table1_KeyUp(ByVal KeyCode)
	If KeyUpHandler(keycode) Then Exit Sub
	If keycode = PlungerKey Then Plunger.Fire:PlaySoundAtVol"plunger", Plunger, 1
End Sub

'**********************************************************************************************************

 ' Drain hole
Sub Drain_Hit:playsoundAtVol"drain", drain, 1:bsTrough.addball me:End Sub

 ' Bumpers
 Sub Bumper1_Hit:vpmTimer.PulseSw 10 : playsoundAtVol"fx_bumper1", Bumper1, VolBump: End Sub
 Sub Bumper2_Hit:vpmTimer.PulseSw 10 : playsoundAtVol"fx_bumper1", Bumper2, VolBump: End Sub

 ' Rollovers
  Sub sw11_Hit:Controller.Switch(11) = 1 : playsoundAtVol"rollover",ActiveBall, 1 : End Sub
  Sub sw11_UnHit:Controller.Switch(11) = 0:End Sub
  Sub sw12_Hit:Controller.Switch(12) = 1 : playsoundAtVol"rollover",ActiveBall, 1 : End Sub
  Sub sw12_UnHit:Controller.Switch(12) = 0:End Sub
  Sub sw13_Hit:Controller.Switch(13) = 1 : playsoundAtVol"rollover", ActiveBall, 1 : End Sub
  Sub sw13_UnHit:Controller.Switch(13) = 0:End Sub
  Sub sw30_Hit:Controller.Switch(30) = 1 : playsoundAtVol"rollover", ActiveBall, 1 : End Sub
  Sub sw30_UnHit:Controller.Switch(30) = 0:End Sub

  Sub sw11b_Hit:Controller.Switch(11) = 1 : playsoundAtVol"rollover",ActiveBall, 1 : End Sub
  Sub sw11b_UnHit:Controller.Switch(11) = 0:End Sub
  Sub sw12b_Hit:Controller.Switch(12) = 1 : playsoundAtVol"rollover",ActiveBall, 1 : End Sub
  Sub sw12b_UnHit:Controller.Switch(12) = 0:End Sub
  Sub sw12d_Hit:Controller.Switch(12) = 1 : playsoundAtVol"rollover",ActiveBall, 1 : End Sub
  Sub sw12d_UnHit:Controller.Switch(12) = 0:End Sub
  Sub sw13b_Hit:Controller.Switch(13) = 1 : playsoundAtVol"rollover",ActiveBall, 1 : End Sub
  Sub sw13b_UnHit:Controller.Switch(13) = 0:End Sub

  'Scoring Rubbers
  Sub sw40a_Slingshot():vpmTimer.PulseSw 40:End Sub
  Sub sw40b_Slingshot():vpmTimer.PulseSw 40:End Sub
  Sub sw40c_Slingshot():vpmTimer.PulseSw 40:End Sub
  Sub sw40d_Slingshot():vpmTimer.PulseSw 40:End Sub
  Sub sw40e_Slingshot():vpmTimer.PulseSw 40:End Sub
  Sub sw40f_Slingshot():vpmTimer.PulseSw 40:End Sub
  Sub sw40g_Slingshot():vpmTimer.PulseSw 40:End Sub
  Sub sw40h_Slingshot():vpmTimer.PulseSw 40:End Sub
  Sub sw40i_Slingshot():vpmTimer.PulseSw 40:End Sub
  Sub sw40j_Slingshot():vpmTimer.PulseSw 40:End Sub

  'Drop targets
  Sub sw21_Hit:dtjj.Hit 1:End Sub
  Sub sw22_Hit:dtjj.Hit 2:End Sub
  Sub sw23_Hit:dtjj.Hit 3:End Sub

  Sub sw31_Hit:dtqq.Hit 1:End Sub
  Sub sw32_Hit:dtqq.Hit 2:End Sub
  Sub sw33_Hit:dtqq.Hit 3:End Sub

  Sub sw41_Hit:dtkk.Hit 1:End Sub
  Sub sw42_Hit:dtkk.Hit 2:End Sub
  Sub sw43_Hit:dtkk.Hit 3:End Sub
  Sub sw44_Hit:dtkk.Hit 4:End Sub

  Sub sw50_Hit:dtaa.Hit 1:End Sub
  Sub sw51_Hit:dtaa.Hit 2:End Sub
  Sub sw52_Hit:dtaa.Hit 3:End Sub
  Sub sw53_Hit:dtaa.Hit 4:End Sub
  Sub sw54_Hit:dtaa.Hit 5:End Sub

  'Standup target
  Sub sw20_Hit:vpmTimer.PulseSw 20:End Sub

  'Drop Target Reset with lamp
  Dim N1,O1
  N1=0:O1=0
  Set LampCallback=GetRef("UpdateMultipleLamps")
  Sub UpdateMultipleLamps
 	N1=Controller.Lamp(17)
    If N1<>O1 Then
		If N1 Then dtkk.DropSol_On
	O1=N1
	End If
  End Sub


Set Lights(4)=l4
Set Lights(8)=l8
Set Lights(9)=l9
Set Lights(10)=l10
Set Lights(11)=l11
Set Lights(12)=l12
Set Lights(13)=l13
Set Lights(14)=l14
Set Lights(15)=l15
Set Lights(16)=l16
Set Lights(18)=l18
Set Lights(19)=l19
Set Lights(20)=l20
Set Lights(21)=l21
Set Lights(22)=l22
Set Lights(23)=l23
Set Lights(24)=l24
Set Lights(25)=l25
Set Lights(26)=l26
Set Lights(27)=l27
Set Lights(28)=l28

 '**********************************************************************************************************
' Backglass Light Displays (7 digit 7 segment displays)
Dim Digits(28)

Digits(0)=Array(a00,a01,a02,a03,a04,a05,a06,n,a08)
Digits(1)=Array(a10,a11,a12,a13,a14,a15,a16,n,a18)
Digits(2)=Array(a20,a21,a22,a23,a24,a25,a26,n,a28)
Digits(3)=Array(a30,a31,a32,a33,a34,a35,a36,n,a38)
Digits(4)=Array(a40,a41,a42,a43,a44,a45,a46,n,a48)
Digits(5)=Array(a50,a51,a52,a53,a54,a55,a56,n,a58)

Digits(6)=Array(b20,b21,b22,b23,b24,b25,b26,n,b28)
Digits(7)=Array(b30,b31,b32,b33,b34,b35,b36,n,b38)
Digits(8)=Array(b40,b41,b42,b43,b44,b45,b46,n,b48)
Digits(9)=Array(b50,b51,b52,b53,b54,b55,b56,n,b58)
Digits(10)=Array(c00,c01,c02,c03,c04,c05,c06,n,c08)
Digits(11)=Array(c10,c11,c12,c13,c14,c15,c16,n,c18)

Digits(12)=Array(c20,c21,c22,c23,c24,c25,c26,n,c28)
Digits(13)=Array(c30,c31,c32,c33,c34,c35,c36,n,c38)
Digits(14)=Array(c40,c41,c42,c43,c44,c45,c46,n,c48)
Digits(15)=Array(c50,c51,c52,c53,c54,c55,c56,n,c58)
Digits(16)=Array(d00,d01,d02,d03,d04,d05,d06,n,d08)
Digits(17)=Array(d10,d11,d12,d13,d14,d15,d16,n,d18)

Digits(18)=Array(d40,d41,d42,d43,d44,d45,d46,n,d48)
Digits(19)=Array(d50,d51,d52,d53,d54,d55,d56,n,d58)
Digits(20)=Array(e00,e01,e02,e03,e04,e05,e06,n,e08)
Digits(21)=Array(e10,e11,e12,e13,e14,e15,e16,n,e18)
Digits(22)=Array(f00,f01,f02,f03,f04,f05,f06,n,f08)
Digits(23)=Array(f10,f11,f12,f13,f14,f15,f16,n,f18)

'credit -- Ball In Play
Digits(24) = Array(e2,e3,e7,e4,e5,e1,e6,n,e23)
Digits(25) = Array(e9,e17,e22,e19,e20,e8,e21,n,e24)
Digits(26) = Array(f2,f3,f7,f4,f5,f1,f6,n,f23)
Digits(27) = Array(f9,f17,f22,f19,f20,f8,f21,n,f24)

Sub DisplayTimer_Timer
	Dim ChgLED,ii,num,chg,stat,obj
	ChgLed = Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
If Not IsEmpty(ChgLED) Then
		If DesktopMode = True Then
		For ii = 0 To UBound(chgLED)
			num = chgLED(ii, 0) : chg = chgLED(ii, 1) : stat = chgLED(ii, 2)
			if (num < 28) then
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
'**********************************************************************************************************
'**********************************************************************************************************
  'Gottlieb System 1
  'added by Inkochnito
  Sub editDips
 	Dim vpmDips : Set vpmDips = New cvpmDips
 	With vpmDips
 		.AddForm 700,400,"System 1 - DIP switches"
 		.AddFrame 205,0,190,"Maximum credits",&H00030000,Array("5 credits",0,"8 credits",&H00020000,"10 credits",&H00010000,"15 credits",&H00030000)'dip 17&18
 		.AddFrame 0,0,190,"Coin chute control",&H00040000,Array("seperate",0,"same",&H00040000)'dip 19
 		.AddFrame 0,46,190,"Game mode",&H00000400,Array("extra ball",0,"replay",&H00000400)'dip 11
 		.AddFrame 0,92,190,"High game to date awards",&H00200000,Array("no award",0,"3 replays",&H00200000)'dip 22
 		.AddFrame 0,138,190,"Balls per game",&H00000100,Array("5 balls",0,"3 balls",&H00000100)'dip 9
 		.AddFrame 0,184,190,"Tilt effect",&H00000800,Array("game over",0,"ball in play only",&H00000800)'dip 12
 		.AddChk 205,80,190,Array("Match feature",&H00000200)'dip 10
 		.AddChk 205,95,190,Array("Credits displayed",&H00001000)'dip 13
 		.AddChk 205,110,190,Array("Play credit button tune",&H00002000)'dip 14
 		.AddChk 205,125,190,Array("Play tones when scoring",&H00080000)'dip 20
 		.AddChk 205,140,190,Array("Play coin switch tune",&H00400000)'dip 23
 		.AddChk 205,155,190,Array("High game to date displayed",&H00100000)'dip 21
 		.AddLabel 50,240,300,20,"After hitting OK, press F3 to reset game with new settings."
 		.ViewDips
 	End With
  End Sub
  Set vpmShowDips = GetRef("editDips")
'**********************************************************************************************************
'**********************************************************************************************************





'VPX callbacks
'**********************************************************************************************************
'**********************************************************************************************************

'*****GI Lights On
dim xx
For each xx in GI:xx.State = 1: Next

'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim Lstep

Sub LeftSlingShot_Slingshot
	vpmTimer.PulseSw 40
    PlaySoundAtVol "right_slingshot", sling2, 1
    LSling.Visible = 0
    LSling1.Visible = 1
    sling2.TransZ = -20
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
		PlaySound "fx_rubber2", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End if
	If finalspeed >= 6 AND finalspeed <= 20 then
 		RandomSoundRubber()
 	End If
End Sub

Sub Posts_Hit(idx)
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 16 then
		PlaySound "fx_rubber2", 0, Vol(ActiveBall)*VolPo, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End if
	If finalspeed >= 6 AND finalspeed <= 16 then
 		RandomSoundRubber()
 	End If
End Sub

Sub RandomSoundRubber()
	Select Case Int(Rnd*3)+1
		Case 1 : PlaySound "rubber_hit_1", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 2 : PlaySound "rubber_hit_2", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 3 : PlaySound "rubber_hit_3", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
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
		Case 1 : PlaySound "flip_hit_1", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 2 : PlaySound "flip_hit_2", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 3 : PlaySound "flip_hit_3", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End Select
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

Sub PlaySoundAtVol(sound, tableobj, Volume)
  PlaySound sound, 1, Volume, Pan(tableobj), 0,0,0, 1, AudioFade(tableobj)
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
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / (VolDiv/VolCol), Pan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub

' Thalamus : Exit in a clean and proper way
Sub Table1_exit()
  Controller.Pause = False
  Controller.Stop
End Sub


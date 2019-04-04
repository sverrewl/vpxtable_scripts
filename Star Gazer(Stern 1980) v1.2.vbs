Option Explicit
Randomize

' Thalamus 2019 March : Improved directional sounds
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

Dim DesktopMode: DesktopMode = Table1.ShowDT

LoadVPM "01560000", "stern.VBS", 3.36


 '******************* Options *********************
' DMD/Backglass Controller Setting
Const cController = 0		'0=Use value defined in cController.txt, 1=VPinMAME, 2=UVP server, 3=B2S server, 4=B2S with DOF (disable VP mech sounds)
'*************************************************

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
'Toggle DOF sounds on/off based on cController value
'*************************************************************
Dim ToggleMechSounds
Function SoundFX (sound)
    If cNewController= 4 and ToggleMechSounds = 0 Then
        SoundFX = ""
    Else
        SoundFX = sound
    End If
End Function

Sub DOF(dofevent, dofstate)
	If cNewController>2 Then
		If dofstate = 2 Then
			Controller.B2SSetData dofevent, 1:Controller.B2SSetData dofevent, 0
		Else
			Controller.B2SSetData dofevent, dofstate
		End If
	End If
End Sub


'***********************************************
'***********************************************
'Standard definitions
'***********************************************
'***********************************************

	 Const cGameName = "stargzr"
     Const UseSolenoids = 2
     Const UseLamps = True
     Const UseSync = 1
     Const UseGI=0
     Const HandleMech = 0
     Const SSolenoidOn = "Solenoid"
     Const SSolenoidOff = ""
     Const SFlipperOn    = "fx_FlipperUp"
     Const SFlipperOff   = "fx_FlipperDown"
     Const SCoin = "CoinIn"
     Const cCredits = "Stern Star Gazer"

'***********************************************
'***********************************************
' Table init.
'***********************************************
'***********************************************

    Dim xx
    Dim bsTrough,dtL,dtR,dtC


  Sub Table1_Init
    vpmInit me
	With Controller
		.GameName = cGameName
		If Err Then MsgBox "Can't start Game " & cGameName & vbNewLine & Err.Description:Exit Sub
		.SplashInfoLine = cCredits
		.HandleKeyboard = 0
		.ShowTitle = 0
		.ShowDMDOnly = 1
		.ShowFrame = 0
		.HandleMechanics = 0
		.Hidden = 1
		On Error Resume Next
		If Err Then MsgBox Err.Description
	End With
    On Error Goto 0
    Controller.Run

'Trough
    Set bsTrough = New cvpmBallStack
    With bsTrough
     .InitNoTrough ballrelease, 33, 90, 4
     .InitExitSnd SoundFX("ballrelease"), SoundFX("Solenoid")
     .BallImage="ball"
    End With

'Nudging
    	vpmNudge.TiltSwitch = 7
    	vpmNudge.Sensitivity= 5
    	vpmNudge.TiltObj=Array(Bumper1,Bumper2,Bumper3,LeftSlingShot,RightSlingshot)

'DropTargets
	Set dtL=New cvpmDropTarget
		dtL.InitDrop Array(sw22,sw23,sw24),Array(22,23,24)
		dtL.InitSnd SoundFX("DTL"), SoundFX("dtResetL")


	Set dtR=New cvpmDropTarget
		dtR.InitDrop Array(sw28,sw29,sw30),Array(28,29,30)
		dtR.InitSnd SoundFX("DTR"), SoundFX("dtrResetR")

	Set dtC=New cvpmDropTarget
		dtC.InitDrop Array(sw25,sw26,sw27),Array(25,26,27)
		dtC.InitSnd SoundFX("DTC"), SoundFX("dtLReset")


    if Controller.Dip(0) <1 then Controller.Pause = true:msgBox "Please select your game options!" & vbNewLine & "Select Disable Start Menu to disable menu.":vpmShowDips:Controller.Pause = false

'Main Timer init
	PinMAMETimer.Interval = PinMAMEInterval
	PinMAMETimer.Enabled = 1

 'GI Init
   VpmMapLights AllLights

'*****GI Lights On
    For each xx in GI:xx.State = 1: Next
    LSling1.Visible = False:LSling2.Visible = False
    RSling1.Visible = False:RSling2.Visible = False
End Sub

Sub Table1_Paused:Controller.Pause = 1:End Sub
Sub Table1_unPaused:Controller.Pause = 0:End Sub

dim dc
dc =1

Sub Table1_KeyDown(ByVal keycode)
    If vpmKeyDown(keycode) Then Exit Sub
	If keycode = PlungerKey Then
		Plunger.PullBack
		PlaySoundAtVol "plungerpull", Plunger, 1
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
End Sub

Sub Table1_KeyUp(ByVal keycode)
    If vpmKeyUp(keycode) Then Exit Sub
	If keycode = PlungerKey Then
		Plunger.Fire
		PlaySoundAtVol "plunger", Plunger, 1
	End If
End Sub

'Flipper Subs
SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

'Solenoids
Const sLSling=1
Const sRSling=2
Const sdtM=3
Const sdtR=4
Const sTopJet=5
Const sKnocker=6
Const sdtL=7
Const sRightJet=8

Const sLeftJet=11
Const sBallRelease=12

SolCallback(sLSling)="vpmSolSound SoundFX(""left_slingshot""),"
SolCallback(sRSling)="vpmSolSound SoundFX(""right_slingshot""),"
SolCallback(sdtM)="SolDropTargetsCenter"
SolCallback(sdtR)="SolDropTargetsRight"
SolCallback(sTopJet)="vpmSolSound SoundFX(""bumper1""),"
SolCallback(sKnocker)="vpmSolSound SoundFX(""knocker""),"
SolCallback(sdtL)="SolDropTargetsLeft"
SolCallback(sRightJet)="vpmSolSound SoundFX(""bumper2""),"
SolCallback(sLeftJet)="vpmSolSound SoundFX(""bumper3""),"
SolCallback(sBallRelease)="bsTrough.SolOut"


Sub SolLFlipper(Enabled)
   If Enabled Then
       PlaySoundAtVol SoundFX("fx_flipperup"), LeftFlipper, VoLFlip:LeftFlipper.RotateToEnd
   Else
       PlaySoundAtVol SoundFX("fx_flipperdown"), LeftFlipper, VolFlip:LeftFlipper.RotateToStart
   End If
End Sub

Sub SolRFlipper(Enabled)
   If Enabled Then
       PlaySoundAtVol SoundFX("fx_flipperup"), RightFlipper, VolFlip:RightFlipper.RotateToEnd
   Else
       PlaySoundAtVol SoundFX("fx_flipperdown"), RightFlipper, VolFlip:RightFlipper.RotateToStart
   End If
End Sub

Sub SolDropTargetsRight(enabled)
	if enabled then
		dtR.DropSol_On
		sw28p.transy = -10
        sw29p.transy = -10
        sw30p.transy = -10
	end if
End Sub

Sub SolDropTargetsLeft(enabled)
	if enabled then
        dtL.DropSol_On
		sw22p.transy = -10
        sw23p.transy = -10
        sw24p.transy = -10
    end if
End Sub

Sub SolDropTargetsCenter(enabled)
	if enabled then
        dtC.DropSol_On
		sw25p.transy = -10
        sw26p.transy = -10
        sw27p.transy = -10
    end if
End Sub

Sub Drain_Hit()
	PlaySoundAtVol "drain", drain, 1
	bsTrough.AddBall Me
End Sub

  'Spinner
Sub sw4_Spin():vpmTimer.pulsesw 4:PlaySoundAtVol "fx_spinner",sw4,VolSpin:End Sub
Sub sw5_Spin():vpmTimer.pulsesw 5:PlaySoundAtVol "fx_spinner",sw5,VolSpin:End Sub
Sub sw9_Spin():vpmTimer.pulsesw 9:PlaySoundAtVol "fx_spinner",sw9,VolSpin:End Sub


Sub Bumper1_Hit
	PlaySoundAtVol "fx_bumper1", ActiveBall, VolBump: vpmTimer.pulsesw(12)
	B1L1.State = 1:B1L2. State = 1
	Me.TimerEnabled = 1
End Sub

Sub Bumper1_Timer
	B1L1.State = 0:B1L2. State = 0
	Me.Timerenabled = 0
End Sub

Sub Bumper2_Hit
	PlaySoundAtVol "fx_bumper2", ActiveBall, VolBump: vpmTimer.pulsesw(13)
	B2L1.State = 1:B2L2. State = 1
	Me.TimerEnabled = 1
End Sub

Sub Bumper2_Timer
	B2L1.State = 0:B2L2. State = 0
	Me.Timerenabled = 0
End Sub

Sub Bumper3_Hit
	PlaySoundAtVol "fx_bumper3", ActiveBall, VolBump: vpmTimer.pulsesw(14)
	B3L1.State = 1:B3L2. State = 1
	Me.TimerEnabled = 1
End Sub

Sub Bumper3_Timer
	B3L1.State = 0:B3L2. State = 0
	Me.Timerenabled = 0
End Sub

'****Drop Targets
Sub sw22_hit
    if sw22p.transy=-10 then
      sw22p.transy = -20:sw22.TimerEnabled=True
    end if
End Sub
Sub sw22_timer
  sw22p.transy = sw22p.transy - 5
  if sw22p.transy = -60 then dtL.Hit 1:vpmTimer.pulsesw(22):sw22.TimerEnabled=False
End Sub

Sub sw23_hit
    if sw23p.transy=-10 then
      sw23p.transy = -20:sw23.TimerEnabled=True
    end if
End Sub
Sub sw23_timer
  sw23p.transy = sw23p.transy - 5
  if sw23p.transy = -60 then dtL.Hit 2:vpmTimer.pulsesw(23):sw23.TimerEnabled=False
End Sub

Sub sw24_hit
    if sw24p.transy=-10 then
      sw24p.transy = -20:sw24.TimerEnabled=True
    end if
End Sub
Sub sw24_timer
  sw24p.transy = sw24p.transy - 5
  if sw24p.transy = -60 then dtL.Hit 3:vpmTimer.pulsesw(24):sw24.TimerEnabled=False
End Sub

Sub sw25_hit
    if sw25p.transy=-10 then
      sw25p.transy = -20:sw25.TimerEnabled=True
    end if
End Sub
Sub sw25_timer
  sw25p.transy = sw25p.transy - 5
  if sw25p.transy = -60 then dtC.Hit 1:vpmTimer.pulsesw(25):sw25.TimerEnabled=False
End Sub

Sub sw26_hit
    if sw26p.transy=-10 then
      sw26p.transy = -20:sw26.TimerEnabled=True
    end if
End Sub
Sub sw26_timer
  sw26p.transy = sw26p.transy - 5
  if sw26p.transy = -60 then dtC.Hit 2:vpmTimer.pulsesw(26):sw26.TimerEnabled=False
End Sub

Sub sw27_hit
    if sw27p.transy=-10 then
      sw27p.transy = -20:sw27.TimerEnabled=True
    end if
End Sub
Sub sw27_timer
  sw27p.transy = sw27p.transy - 5
  if sw27p.transy = -60 then dtC.Hit 3:vpmTimer.pulsesw(27):sw27.TimerEnabled=False
End Sub

Sub sw28_hit
    if sw28p.transy=-10 then
      sw28p.transy = -20:sw28.TimerEnabled=True
    end if
End Sub
Sub sw28_timer
  sw28p.transy = sw28p.transy - 5
  if sw28p.transy = -60 then dtR.Hit 1::vpmTimer.pulsesw(28):sw28.TimerEnabled=False
End Sub

Sub sw29_hit
    if sw29p.transy=-10 then
      sw29p.transy = -20:sw29.TimerEnabled=True
    end if
End Sub
Sub sw29_timer
  sw29p.transy = sw29p.transy - 5
  if sw29p.transy = -60 then dtR.Hit 2:vpmTimer.pulsesw(29):sw29.TimerEnabled=False
End Sub

Sub sw30_hit
    if sw30p.transy=-10 then
      sw30p.transy = -20:sw30.TimerEnabled=True
    end if
End Sub
Sub sw30_timer
  sw30p.transy = sw30p.transy - 5
  if sw30p.transy = -60 then dtR.Hit 3:vpmTimer.pulsesw(30):sw30.TimerEnabled=False
End Sub

' ** Square Targets
Sub sw10_hit
  vpmTimer.pulsesw(10):sw10p.transX=10:sw10.TimerEnabled=True
End Sub
Sub sw10_timer
  me.TimerEnabled=False:sw10p.transX=0
End Sub

Sub sw11_hit
  vpmTimer.pulsesw(11):sw11p.transX=10:sw11.TimerEnabled=True
End Sub
Sub sw11_timer
  me.TimerEnabled=False:sw11p.transX=0
End Sub

Sub sw17_hit
  vpmTimer.pulsesw(17):sw17p.transX=10:sw17.TimerEnabled=True
End Sub
Sub sw17_timer
  me.TimerEnabled=False:sw17p.transX=0
End Sub

Sub sw18_hit
  vpmTimer.pulsesw(18):sw18p.transX=10:sw18.TimerEnabled=True
End Sub
Sub sw18_timer
  me.TimerEnabled=False:sw18p.transX=0
End Sub

Sub sw19_hit
  vpmTimer.pulsesw(19):sw19p.transX=10:sw19.TimerEnabled=True
End Sub
Sub sw19_timer
  me.TimerEnabled=False:sw19p.transX=0
End Sub

Sub sw20_hit
  vpmTimer.pulsesw(20):sw20p.transX=10:sw20.TimerEnabled=True
End Sub
Sub sw20_timer
  me.TimerEnabled=False:sw20p.transX=0
End Sub

Sub sw21_hit
  vpmTimer.pulsesw(21):sw21p.transX=10:sw21.TimerEnabled=True
End Sub
Sub sw21_timer
  me.TimerEnabled=False:sw21p.transX=0
End Sub

Sub sw31_hit
  vpmTimer.pulsesw(31):sw31p.transX=10:sw31.TimerEnabled=True
End Sub
Sub sw31_timer
  me.TimerEnabled=False:sw31p.transX=0
End Sub

Sub sw32_hit
  vpmTimer.pulsesw(32):sw32p.transX=10:sw32.TimerEnabled=True
End Sub
Sub sw32_timer
  me.TimerEnabled=False:sw32p.transX=0
End Sub

Sub sw38_hit
  vpmTimer.pulsesw(38):sw38p.transX=10:sw38.TimerEnabled=True
End Sub
Sub sw38_timer
  me.TimerEnabled=False:sw38p.transX=0
End Sub

Sub sw39_hit
  vpmTimer.pulsesw(39):sw39p.transX=10:sw39.TimerEnabled=True
End Sub
Sub sw39_timer
  me.TimerEnabled=False:sw39p.transX=0
End Sub

Sub sw40_hit
  vpmTimer.pulsesw(40):sw40p.transX=10:sw40.TimerEnabled=True
End Sub
Sub sw40_timer
  me.TimerEnabled=False:sw40p.transX=0
End Sub

'Rollovers
Sub sw34_hit:Controller.Switch(34)=1:End Sub
Sub sw34_unhit:Controller.Switch(34)=0:End Sub
Sub sw35_hit:Controller.Switch(35)=1:End Sub
Sub sw35_unhit:Controller.Switch(35)=0:End Sub
Sub sw36_hit:Controller.Switch(36)=1:End Sub
Sub sw36_unhit:Controller.Switch(36)=0:End Sub
Sub sw37_hit:Controller.Switch(37)=1:End Sub
Sub sw37_unhit:Controller.Switch(37)=0:End Sub

' RealTimeUpdates
Set MotorCallback = GetRef("GameTimer")

Sub GameTimer
	UpdateGates
End Sub

Sub UpdateGates
	Gate3D.RotX = -Gate1.CurrentAngle/1.5-30
End Sub


'**********************************************************************************************************
' Backglass Light Displays (7 digit 7 segment displays)
Dim Digits(32)

' 1st Player
Digits(0) = Array(LED10,LED11,LED12,LED13,LED14,LED15,LED16)
Digits(1) = Array(LED20,LED21,LED22,LED23,LED24,LED25,LED26)
Digits(2) = Array(LED30,LED31,LED32,LED33,LED34,LED35,LED36)
Digits(3) = Array(LED40,LED41,LED42,LED43,LED44,LED45,LED46)
Digits(4) = Array(LED50,LED51,LED52,LED53,LED54,LED55,LED56)
Digits(5) = Array(LED60,LED61,LED62,LED63,LED64,LED65,LED66)
Digits(6) = Array(LED70,LED71,LED72,LED73,LED74,LED75,LED76)

' 2nd Player
Digits(7) = Array(LED80,LED81,LED82,LED83,LED84,LED85,LED86)
Digits(8) = Array(LED90,LED91,LED92,LED93,LED94,LED95,LED96)
Digits(9) = Array(LED100,LED101,LED102,LED103,LED104,LED105,LED106)
Digits(10) = Array(LED110,LED111,LED112,LED113,LED114,LED115,LED116)
Digits(11) = Array(LED120,LED121,LED122,LED123,LED124,LED125,LED126)
Digits(12) = Array(LED130,LED131,LED132,LED133,LED134,LED135,LED136)
Digits(13) = Array(LED140,LED141,LED142,LED143,LED144,LED145,LED146)

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

' Credits
Digits(28) = Array(LED4,LED2,LED6,LED7,LED5,LED1,LED3)
Digits(29) = Array(LED18,LED9,LED27,LED28,LED19,LED8,LED17)
' Balls
Digits(30) = Array(LED39,LED37,LED48,LED49,LED47,LED29,LED38)
Digits(31) = Array(LED67,LED58,LED69,LED77,LED68,LED57,LED59)

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
				if char(stat) > "" then msg(num) = char(stat)
			end if
		next
		end if
end if
End Sub




'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep

Sub RightSlingShot_Slingshot
    PlaySoundAtVol "right_slingshot", sling1, 1
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.TransZ = -20
    RStep = 0
    RightSlingShot.TimerEnabled = 1
    vpmTimer.pulsesw(15)
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.TransZ = -10
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.TransZ = 0:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
    PlaySoundAtVol "left_slingshot", sling2, 1
    LSling.Visible = 0
    LSling1.Visible = 1
    sling2.TransZ = -20
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
    vpmTimer.pulsesw(16)
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.TransZ = -10
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.TransZ = 0:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

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


' '*****************************************
' '    JP's VP10 Collision & Rolling Sounds
' '*****************************************
'
' Const tnob = 2 ' total number of balls
' ReDim rolling(tnob)
' ReDim collision(tnob)
' Initcollision
'
' Sub Initcollision
'     Dim i
'     For i = 0 to tnob
'         collision(i) = -1
'         rolling(i) = False
'     Next
' End Sub
'
' Sub CollisionTimer_Timer()
'     Dim BOT, B, B1, B2, dx, dy, dz, distance, radii
'     BOT = GetBalls
'
'     ' rolling
'
' 	For B = UBound(BOT) +1 to tnob
'         rolling(b) = False
'         StopSound("fx_ballrolling" & b)
' 	Next
'
'     If UBound(BOT) = -1 Then Exit Sub
'
'     For B = 0 to UBound(BOT)
'         If BallVel(BOT(b) ) > 1 AND BOT(b).z < 30 Then
'             rolling(b) = True
'             PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) ), Pan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0
'         Else
'             If rolling(b) = True Then
'                 StopSound("fx_ballrolling" & b)
'                 rolling(b) = False
'             End If
'         End If
'     Next
'
'     'collision
'
'     If UBound(BOT) < 1 Then Exit Sub
'
'     For B1 = 0 to UBound(BOT)
'         For B2 = B1 + 1 to UBound(BOT)
'             dz = INT(ABS((BOT(b1).z - BOT(b2).z) ) )
'             radii = BOT(b1).radius + BOT(b2).radius
' 			If dz <= radii Then
'
'             dx = INT(ABS((BOT(b1).x - BOT(b2).x) ) )
'             dy = INT(ABS((BOT(b1).y - BOT(b2).y) ) )
'             distance = INT(SQR(dx ^2 + dy ^2) )
'
'             If distance <= radii AND (collision(b1) = -1 OR collision(b2) = -1) Then
'                 collision(b1) = b2
'                 collision(b2) = b1
'                 PlaySound("fx_collide"), 0, Vol2(BOT(b1), BOT(b2)), Pan(BOT(b1)), 0, Pitch(BOT(b1)), 0, 0
'             Else
'                 If distance > (radii + 10)  Then
'                     If collision(b1) = b2 Then collision(b1) = -1
'                     If collision(b2) = b1 Then collision(b2) = -1
'                 End If
'             End If
' 			End If
'         Next
'     Next
' End Sub
'

'*****************************************
'      JP's VP10 Rolling Sounds
'*****************************************

Const tnob = 2 ' total number of balls
ReDim rolling(tnob)
InitRolling

Sub InitRolling
    Dim i
    For i = 0 to tnob
        rolling(i) = False
    Next
End Sub

' Sub RollingTimer_Timer()
Sub CollisionTimer_Timer()
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
		PlaySound "fx_rubber2", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
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

'Stern Star Gazer
'added by Inkochnito
Sub editDips
	Dim vpmDips:Set vpmDips=New cvpmDips
	With vpmDips
		.AddForm 700,400,"Star Gazer - DIP switches"
		.AddChk 2,10,180,Array("Match feature",&H00100000)'dip 21
		.AddChk 2,25,180,Array("Credits display",&H00080000)'dip 20
		.AddChk 2,40,195,Array("Bottom banks spot next zodiac target",&H10000000)'dip 29
		.AddChk 2,55,180,Array("Extra ball awarded",&H20000000)'dip 30
		.AddChk 2,70,180,Array("Background sound (Unavailable?)",&H00000080)'dip 8
		.AddFrame 2,90,190,"Maximum credits",&H00060000,Array("10 credits",0,"15 credits",&H00020000,"25 credits",&H00040000,"40 credits",&H00060000)'dip 18&19
		.AddFrame 2,166,190,"High game to date",49152,Array("points",0,"1 free game",&H00004000,"2 free games",32768,"3 free games",49152)'dip 15&16
		.AddFrame 2,242,190,"Special award",&HC0000000,Array("no award",0,"100,000 points",&H40000000,"free ball",&H80000000,"free game",&HC0000000)'dip 31&32
		.AddFrame 205,10,190,"Zodiac special when ring complete",&H00A00000,Array("2nd time",0,"3rd time",&H00800000,"4th time",&H00A00000)'dip 22&24
		.AddFrame 205,71,190,"Add-a-ball memory",&H00003000,Array("one only",0,"maximum three",&H00001000,"maximum five",&H00003000)'dip 13&14
		.AddFrame 205,132,190,"High score feature",&H00000020,Array("extra ball",0,"replay",&H00000020)'dip 6
		.AddFrame 205,178,190,"Balls per game",&H00000040,Array("3 balls",0,"5 balls",&H00000040)'dip 7
		.AddFrame 205,224,190,"Special adjust",&H00400000,Array("target bank on, outlanes alternating",0,"all 3 positions alternate",&H00400000)'dip 23
		.AddFrame 205,270,190,"Flashing lite speed",&H00000010,Array("fast",0,"slow",&H00000010)'dip 5
		.AddFrame 100,316,190,"Extra ball lite",&H00010000,Array("alternate",0,"stays on",&H00010000)'dip 17
        .AddFrame 100,365,190,"Menu at Start",&H00000001,Array("Enable",0,"Disable", &H00000001) ' dip1
		.AddLabel 50,420,300,20,"After hitting OK, press F3 to reset game with new settings."
        .AddLabel 50,440,300,20,"Press F6 anytime to return to this menu"
		.ViewDips
	End With
End Sub
Set vpmShowDips=GetRef("editDips")

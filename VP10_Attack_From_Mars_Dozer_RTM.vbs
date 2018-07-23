' Attackfrom Mars
' Based on the table by Bally/Williams
' VP916 3.0 by JPSalas 2013

Option Explicit
Randomize

Dim DesktopMode:DesktopMode = Table1.ShowDT
Dim UseVPMDMD

UseVPMDMD = DesktopMode

Dim LightHalo_ON, SpinningSaucers, BlackLights, InnerCabinetFlash, WireFormFlash, UFOAmbient, UFOReflection, Mars_Mod, BlackLightFlashers, BLWALL

'******************* Options *********************
' DMD/Backglass Controller Setting
Const cController = 3		'0=Use value defined in cController.txt, 1=VPinMAME, 2=UVP server, 3=B2S server, 4=B2S with DOF (disable VP mech sounds)
'*************************************************

SpinningSaucers = 1 'Can set to 0 to disable Spinning Saucers...  performance

BlackLights = 1     'Change to 1 to add blacklight saucers Images

BlackLightFlashers = 1 'Turn on the black light flasher effects.

BLWALL = 1 'Illuminate the back cabinet wall and side walls with Black Light Effect

WireFormFlash = 1 ' Swap the Wireform materials to mirror the flashers.

UfoAmbient = 1 ' Render some flasher ambience on the Main Saucer.

UfoTilt = 1 ' Exaggerate the main saucer shake when the game is nudged.

UFOReflection = 1 'Turn on or off the main saucer reflection.

Mars_Mod = 1 'Rotate a fantasy based light mod beneath the saucer when it is hit during attack phase.

'TEMPORARILY OUT OF ACTION > 'InnerCabinetFlash = 1 'Swap the outbox primitive image to provide some flasher ambience on the cab walls.

'**************************************

Const BallSize = 53

'LoadVPM "01560000", "WPC.VBS", 3.26
LoadVPM "01560000", "WPC.VBS", 3.46


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

'********************
'Standard definitions
'********************

Const UseSolenoids = 1
Const UseLamps = 0
Const UseSync = 0
Const HandleMech = 0

' Standard Sounds
Const SSolenoidOn = "fx_Solenoidon"
Const SSolenoidOff = "fx_solenoidoff"
Const SCoin = "fx_Coin"

Set GiCallback2 = GetRef("UpdateGI")

Dim bsTrough, bsL, bsR, dtDrop, x, bump1, bump2, bump3, BallFrame, plungerIM, Mech3bank,DiverterDir, DiverterPos, UfoTilt

'************
' Table init.
'************

Const cGameName = "afm_113b" 'arcade rom - with credits
'Const cGameName = "afm_113"  'home rom - free play

Sub Table1_Init
    vpmInit Me
    With Controller
        .GameName = cGameName
        If Err Then MsgBox "Can't start Game " & cGameName & vbNewLine & Err.Description:Exit Sub
        .SplashInfoLine = "Attack from Mars" & vbNewLine & "VP916 table by JPSalas v3.0"
        'DMD example: position and size for 1400x1050
        '.Games(cGameName).Settings.Value("dmd_pos_x")=500
        '.Games(cGameName).Settings.Value("dmd_pos_y")=2
        '.Games(cGameName).Settings.Value("dmd_width")=400
        '.Games(cGameName).Settings.Value("dmd_height")=92
        .Games(cGameName).Settings.Value("rol") = 0
        .HandleKeyboard = 0
        .ShowTitle = 0
        .ShowDMDOnly = 1
        .ShowFrame = 0
        .HandleMechanics = 0
        .Hidden = DeskTopMode
        On Error Resume Next
        .Run GetPlayerHWnd
        If Err Then MsgBox Err.Description
        On Error Goto 0
        .Switch(22) = 1 'close coin door
        .Switch(24) = 1 'and keep it close
    End With


	StartLampTimer
	SetBlacklights
    BWON.isdropped = 1
    BWONL.isdropped = 1
    BWONM.isdropped = 1
    BWONR.isdropped = 1

    ' Nudging
    vpmNudge.TiltSwitch = 14
    vpmNudge.Sensitivity = 1
    vpmNudge.TiltObj = Array(bumper1, bumper2, bumper3, LeftSlingshot, RightSlingshot)

    ' Trough
    Set bsTrough = New cvpmBallStack
    With bsTrough
        .InitSw 0, 32, 33, 34, 35, 0, 0, 0
        .InitKick BallRelease, 90, 4
        .InitEntrySnd "fx_Solenoid", "fx_Solenoid"
        .InitExitSnd "fx_ballrel", "fx_Solenoid"
        .Balls = 4
    End With

    ' Droptarget
    Set dtDrop = New cvpmDropTarget
    With dtDrop
        .InitDrop sw77, 77
        .initsnd "fx_droptarget", "fx_resetdrop"
    End With

    ' Left hole
    Set bsL = New cvpmBallStack
    With bsL
        .InitSw 0, 36, 0, 0, 0, 0, 0, 0
        .InitKick sw36, 0, 2
        .InitExitSnd "fx_popper", "fx_Solenoid"
        .KickForceVar = 3
    End With

    ' Right hole
    Set bsR = New cvpmBallStack
    With bsR
        .InitSw 0, 37, 0, 0, 0, 0, 0, 0
        .InitKick sw37, 200, 24
        .KickZ = 0.4
        .InitExitSnd "fx_popper", "fx_Solenoid"
        '.KickForceVar = 2
        .KickAngleVar = 2
        .KickBalls = 1
    End With

    '3 Targets Bank
    Set Mech3Bank = new cvpmMech
    With Mech3Bank
        .Sol1 = 24
        .Mtype = vpmMechLinear + vpmMechReverse + vpmMechOneSol
        .Length = 60
        .Steps = 50
        .AddSw 67, 0, 0
        .AddSw 66, 50, 50
        .Callback = GetRef("Update3Bank")
        .Start
    End With

    ' Impulse Plunger
    Const IMPowerSetting = 42 'Plunger Power
    Const IMTime = 0.6        ' Time in seconds for Full Plunge
    Set plungerIM = New cvpmImpulseP
    With plungerIM
        .InitImpulseP swplunger, IMPowerSetting, IMTime
        .Random 0.3
        .switch 18
        .InitExitSnd "fx_plunger", "fx_plunger"
        .CreateEvents "plungerIM"
    End With

    ' Main Timer init
    PinMAMETimer.Interval = PinMAMEInterval
    PinMAMETimer.Enabled = 1
'    StartShake

    ' Init other dropwalls - animations
    UpdateGI 0, 1:UpdateGI 1, 1:UpdateGI 2, 1
    Diverter2.IsDropped = 1:Diverter3.IsDropped = 1:Diverter4.IsDropped = 1:Diverter5.IsDropped = 1:Diverter6.IsDropped = 1
    DiverterPos = 0
    LEDSpeedSlow
    UfoLed.Enabled = 0
'    LeftSLing.IsDropped = 1:LeftSLing2.IsDropped = 1:LeftSLing3.IsDropped = 1
'    RightSLing.IsDropped = 1:RightSLing2.IsDropped = 1:RightSLing3.IsDropped = 1
    LeftSLingH.IsDropped = 1:LeftSLingH2.IsDropped = 1:LeftSLingH3.IsDropped = 1
    RightSLingH.IsDropped = 1:RightSLingH2.IsDropped = 1:RightSLingH3.IsDropped = 1
	sw56a.IsDropped=1:sw57a.IsDropped=1:sw58a.IsDropped=1
	sw75a.IsDropped=1:sw76a.IsDropped=1
	sw41a.IsDropped=1:sw42a.IsDropped=1:sw43a.IsDropped=1:sw44a.IsDropped=1
'	la1.IsDropped=1:la2.IsDropped=1:la3.IsDropped=1:la4.IsDropped=1:la5.IsDropped=1
'	la6.IsDropped=1:la8.IsDropped=1:la9.IsDropped=1:la10.IsDropped=1:la11.IsDropped=1
End Sub

Sub table1_Paused:Controller.Pause = 1:End Sub
Sub table1_unPaused:Controller.Pause = 0:End Sub

 If Table1.ShowDT = true then
Ramp15.visible = 1
Ramp16.visible = 1
FL27UFO.Intensity = 0
FL26UFO.Intensity = 0
FL18UFO.Intensity = 0
FL23.Intensity = 0
FL23a.Intensity = 0
 else
Ramp15.visible = 0
Ramp16.visible = 0
FL27UFO.Intensity = 10
FL26UFO.Intensity = 10
FL18UFO.Intensity = 10
FL23.Intensity = 20
FL23a.Intensity = 20
 End If

Dim bulb, xxlens

If BlackLightFlashers = 1 Then
for each bulb in UFlash
bulb.Color=RGB(128,0,128)
bulb.ColorFull=RGB(128,0,128)
next
for each bulb in UFOAmb
bulb.Color=RGB(128,0,128)
bulb.ColorFull=RGB(128,0,128)
next
for each xxlens in UFO_Flens
xxlens.image="prf_blue"
next
else
for each bulb in UFlash
bulb.Color=RGB(255,0,0)
bulb.ColorFull=RGB(255,0,0)
next
for each bulb in UFOAmb
bulb.Color=RGB(255,0,0)
bulb.ColorFull=RGB(255,0,0)
next
for each xxlens in UFO_Flens
xxlens.image="prf"
next
End If

Dim xxufo
If UFOReflection = 1 Then
For each xxufo in UFOS:xxufo.ReflectionEnabled = 1:next
ufo1d.ReflectionEnabled = 1
else
For each xxufo in UFOS:xxufo.ReflectionEnabled = 0:next
ufo1d.ReflectionEnabled = 0
End If

'**********
' Keys
'**********

Sub table1_KeyDown(ByVal Keycode)
    If keycode = PlungerKey Then Controller.Switch(11) = 1:End If
    If keycode = LeftTiltKey Then LeftNudge 90, 1.6, 20:If UfoTilt = 1 Then smallufoshake:PlaySound "fx_nudge", 0, 1, -0.1, 0.25:End If
    If keycode = RightTiltKey Then RightNudge 270, 1.6, 20:If UfoTilt = 1 Then smallufoshake:PlaySound "fx_nudge", 0, 1, 0.1, 0.25:End If
    If keycode = CenterTiltKey Then CenterNudge 0, 2.8, 30:If UfoTilt = 1 Then smallufoshake:PlaySound "fx_nudge", 0, 1, 0, 0.25:End If
    If vpmKeyDown(keycode) Then Exit Sub
'debug key
    if keycode = "3" then
       Solufoshake 1
    end if
End Sub

Sub table1_KeyUp(ByVal Keycode)
    If keycode = PlungerKey Then Controller.Switch(11) = 0
    If vpmKeyUp(keycode) Then Exit Sub
'debug key
    'if keycode = "3" then
    'end if
End Sub

'*************************************
'          Nudge System
' based on Noah's nudgetest table
'*************************************

Dim LeftNudgeEffect, RightNudgeEffect, NudgeEffect

Sub LeftNudge(angle, strength, delay)
    vpmNudge.DoNudge angle, (strength * (delay-LeftNudgeEffect) / delay) + RightNudgeEffect / delay
    LeftNudgeEffect = delay
    RightNudgeEffect = 0
    RightNudgeTimer.Enabled = 0
    LeftNudgeTimer.Interval = delay
    LeftNudgeTimer.Enabled = 1
End Sub

Sub RightNudge(angle, strength, delay)
    vpmNudge.DoNudge angle, (strength * (delay-RightNudgeEffect) / delay) + LeftNudgeEffect / delay
    RightNudgeEffect = delay
    LeftNudgeEffect = 0
    LeftNudgeTimer.Enabled = 0
    RightNudgeTimer.Interval = delay
    RightNudgeTimer.Enabled = 1
End Sub

Sub CenterNudge(angle, strength, delay)
    vpmNudge.DoNudge angle, strength * (delay-NudgeEffect) / delay
    NudgeEffect = delay
    NudgeTimer.Interval = delay
    NudgeTimer.Enabled = 1
End Sub

Sub LeftNudgeTimer_Timer()
    LeftNudgeEffect = LeftNudgeEffect-1
    If LeftNudgeEffect = 0 then LeftNudgeTimer.Enabled = False
End Sub

Sub RightNudgeTimer_Timer()
    RightNudgeEffect = RightNudgeEffect-1
    If RightNudgeEffect = 0 then RightNudgeTimer.Enabled = False
End Sub

Sub NudgeTimer_Timer()
    NudgeEffect = NudgeEffect-1
    If NudgeEffect = 0 then NudgeTimer.Enabled = False
End Sub

'*********
' Switches
'*********

' Slings & div switches
Dim LStep, RStep

Sub LeftSlingShot_Slingshot:SlingLa.visible = false:SlingLb.visible = true:LeftSLingH.IsDropped = 0:PlaySound "fx_slingshot",0,1,-0.15,0.25:vpmTimer.PulseSw 51:LStep = 0:Me.TimerEnabled = 1:End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 2:SlingLb.visible = false:SlingLc.visible = true:LeftSLingH.IsDropped = 1:LeftSLingH2.IsDropped = 0
        Case 3:SlingLc.visible = false:SlingLd.visible = true:LeftSLingH2.IsDropped = 1:LeftSLingH3.IsDropped = 0
        Case 4:SlingLd.visible = false:SlingLa.visible = true:LeftSLingH3.IsDropped = 1:Me.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

Sub RightSlingShot_Slingshot:SlingRa.visible = false:SlingRb.visible = true:RightSLingH.IsDropped = 0:PlaySound "fx_slingshot",0,1,0.15,0.25:vpmTimer.PulseSw 52:RStep = 0:Me.TimerEnabled = 1:End Sub
Sub RightSlingShot_Timer
    Select Case RStep
        Case 2:SlingRb.visible = false:SlingRc.visible = true:RightSLingH.IsDropped = 1:RightSLingH2.IsDropped = 0
        Case 3:SlingRc.visible = false:SlingRd.visible = true:RightSLingH2.IsDropped = 1:RightSLingH3.IsDropped = 0
        Case 4:SlingRd.visible = false:SlingRa.visible = true:RightSLingH3.IsDropped = 1:Me.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

' Bumpers
Sub Bumper1_Hit:vpmTimer.PulseSw 53:PlaySound "fx_bumper",0,1,0.15,0.25:bump1 = 1:Me.TimerEnabled = 1:End Sub
Sub Bumper1_Timer()
    Select Case bump1
        Case 1:'Ring1.HeightTop = 15:Ring1.HeightBottom = 15:bump1 = 2
        Case 2:'Ring1.HeightTop = 25:Ring1.HeightBottom = 25:bump1 = 3
        Case 3:'Ring1.HeightTop = 35:Ring1.HeightBottom = 35:bump1 = 4
        Case 4:Me.TimerEnabled = 0'Ring1.HeightTop = 45:Ring1.HeightBottom = 45:
    End Select
End Sub

Sub Bumper2_Hit:vpmTimer.PulseSw 54:PlaySound "fx_bumper",0,1,0.15,0.25:bump2 = 1:Me.TimerEnabled = 1:End Sub
Sub Bumper2_Timer()
    Select Case bump2
        Case 1:'Ring2.HeightTop = 15:Ring2.HeightBottom = 15:bump2 = 2
        Case 2:'Ring2.HeightTop = 25:Ring2.HeightBottom = 25:bump2 = 3
        Case 3:'Ring2.HeightTop = 35:Ring2.HeightBottom = 35:bump2 = 4
        Case 4:Me.TimerEnabled = 0'Ring2.HeightTop = 45:Ring2.HeightBottom = 45:
    End Select
End Sub

Sub Bumper3_Hit:vpmTimer.PulseSw 55:PlaySound "fx_bumper",0,1,0.15,0.25:bump3 = 1:Me.TimerEnabled = 1:End Sub
Sub Bumper3_Timer()
    Select Case bump3
        Case 1:'Ring3.HeightTop = 15:Ring3.HeightBottom = 15:bump3 = 2
        Case 2:'Ring3.HeightTop = 25:Ring3.HeightBottom = 25:bump3 = 3
        Case 3:'Ring3.HeightTop = 35:Ring3.HeightBottom = 35:bump3 = 4
        Case 4:Me.TimerEnabled = 0'Ring3.HeightTop = 45:Ring3.HeightBottom = 45:
    End Select
End Sub

' Drain holes, vuks & saucers
Sub Drain_Hit:Playsound "fx_drain":bsTrough.AddBall Me:End Sub
Sub sw36a_Hit:PlaySound "scoopenter",0,1,-0.1,0.25:bsL.AddBall Me:End Sub
Sub sw37a_Hit:PlaySound "fx_kicker_enter",0,1,0.1,0.25:bsR.AddBall Me:End Sub

Dim aBall

Sub sw78_Hit
    PlaySound "fx_ballhit",0,1,0.1,0.25
	Set aBall = ActiveBall:Me.TimerEnabled =1
    
    vpmTimer.PulseSwitch(78), 150, "bsl.addball 0 '"
End Sub

Sub sw78_Timer
	Do While aBall.Z >0
		aBall.Z = aBall.Z -5
		Exit Sub
	Loop
    Me.DestroyBall
	Me.TimerEnabled = 0
End Sub

Dim bBall

Sub sw37_Hit
    PlaySound "fx_ballhit",0,1,0.1,0.25
	Set bBall = ActiveBall:Me.TimerEnabled =1
    
    bsR.AddBall 0
End Sub

Sub sw37_Timer
	Do While bBall.Z >0
		bBall.Z = bBall.Z -5
		Exit Sub
	Loop
    Me.DestroyBall
	Me.TimerEnabled = 0
End Sub


' Rollovers & Ramp Switches
Sub sw16_Hit:Controller.Switch(16) = 1:PlaySound "fx_sensor",0,1,-0.1,0.25:End Sub
Sub sw16_UnHit:Controller.Switch(16) = 0:End Sub

Sub sw26_Hit:Controller.Switch(26) = 1:PlaySound "fx_sensor",0,1,-0.1,0.25:End Sub
Sub sw26_UnHit:Controller.Switch(26) = 0:End Sub

Sub sw17_Hit:Controller.Switch(17) = 1:PlaySound "fx_sensor",0,1,0.1,0.25:End Sub
Sub sw17_UnHit:Controller.Switch(17) = 0:End Sub

Sub sw27_Hit:Controller.Switch(27) = 1:PlaySound "fx_sensor",0,1,0.1,0.25:End Sub
Sub sw27_UnHit:Controller.Switch(27) = 0:End Sub

Sub sw38_Hit:Controller.Switch(38) = 1:PlaySound "fx_sensor",0,1,0.1,0.25:End Sub
Sub sw38_Unhit:Controller.Switch(38) = 0:End Sub

Sub sw48_Hit:Controller.Switch(48) = 1:PlaySound "fx_sensor",0,1,0.1,0.25:End Sub
Sub sw48_Unhit:Controller.Switch(48) = 0:End Sub

Sub sw71_Hit:Controller.Switch(71) = 1:PlaySound "fx_sensor",0,1,0.1,0.25:End Sub
Sub sw71_UnHit:Controller.Switch(71) = 0:End Sub

'Sub sw72_Hit:Controller.Switch(72) = 1:PlaySound "fx_sensor",0,1,0.1,0.25:End Sub
'Sub sw72_Unhit:Controller.Switch(72) = 0:End Sub

Sub sw72_Hit():vpmTimer.PulseSw 72:PlaySound "fx_sensor",0,1,0.1,0.25:End Sub

Sub sw73_Hit:Controller.Switch(73) = 1:PlaySound "fx_sensor",0,1,-0.1,0.25:End Sub
Sub sw73_Unhit:Controller.Switch(73) = 0:End Sub

Sub sw74_Hit:Controller.Switch(74) = 1:PlaySound "fx_sensor",0,1,-0.1,0.25:End Sub
Sub sw74_Unhit:Controller.Switch(74) = 0:End Sub

Sub sw61_Hit:Controller.Switch(61) = 1:End Sub
Sub sw61_Unhit:Controller.Switch(61) = 0:End Sub

Sub sw62_Hit:Controller.Switch(62) = 1:End Sub
Sub sw62_Unhit:Controller.Switch(62) = 0:End Sub

Sub sw63_Hit
    Controller.Switch(63) = 1
    If ActiveBall.VelY <-20 Then ActiveBall.VelY = -20
End Sub

Sub sw63_Unhit:Controller.Switch(63) = 0:End Sub

Sub sw64_Hit:Controller.Switch(64) = 1:PlaySound "fx_metalrolling",0,1,-0.15,0.25:End Sub
Sub sw64_Unhit:Controller.Switch(64) = 0:ActiveBall.VelY= ActiveBall.VelY/3:End Sub

Sub sw65_Hit:Controller.Switch(65) = 1:PlaySound "fx_metalrolling",0,1,0.15,0.25:End Sub
Sub sw65_Unhit:Controller.Switch(65) = 0:End Sub

' Targets

Dim sw44Step, sw43Step, sw42Step, sw41Step, sw58Step, sw57Step, sw56Step, sw45Step, sw46Step, sw47Step, sw76Step, sw75Step

Sub sw56_Hit:vpmTimer.PulseSw 56:swp56.TransX = -5:sw56Step = 0:Me.TimerEnabled = 1:Me.TimerInterval = 10:PlaySound "fx_target",0,1,-0.1,0.25:End Sub
Sub sw56_Timer()
	Select Case sw56Step
		Case 1:swp56.TransX = 3
        Case 2:swp56.TransX = -2
        Case 3:swp56.TransX = 1
        Case 4:swp56.TransX = 0:Me.TimerEnabled = 0
     End Select
	sw56Step = sw56Step + 1
End Sub

Sub sw57_Hit:vpmTimer.PulseSw 57:swp57.TransX = -5:sw57Step = 0:Me.TimerEnabled = 1:Me.TimerInterval = 10:PlaySound "fx_target",0,1,-0.1,0.25:End Sub
Sub sw57_Timer()
	Select Case sw57Step
		Case 1:swp57.TransX = 3
        Case 2:swp57.TransX = -2
        Case 3:swp57.TransX = 1
        Case 4:swp57.TransX = 0:Me.TimerEnabled = 0
     End Select
	sw57Step = sw57Step + 1
End Sub

Sub sw58_Hit:vpmTimer.PulseSw 58:swp58.TransX = -5:sw58Step = 0:Me.TimerEnabled = 1:Me.TimerInterval = 10:PlaySound "fx_target",0,1,-0.1,0.25:End Sub
Sub sw58_Timer()
	Select Case sw58Step
		Case 1:swp58.TransX = 3
        Case 2:swp58.TransX = -2
        Case 3:swp58.TransX = 1
        Case 4:swp58.TransX = 0:Me.TimerEnabled = 0
     End Select
	sw58Step = sw58Step + 1
End Sub

Sub sw41_Hit:vpmTimer.PulseSw 41:swp41.TransX = -5:sw41Step = 0:Me.TimerEnabled = 1:Me.TimerInterval = 10:PlaySound "fx_target",0,1,0.1,0.25:End Sub
Sub sw41_Timer()
	Select Case sw41Step
		Case 1:swp41.TransX = 3
        Case 2:swp41.TransX = -2
        Case 3:swp41.TransX = 1
        Case 4:swp41.TransX = 0:Me.TimerEnabled = 0
     End Select
	sw41Step = sw41Step + 1
End Sub

Sub sw42_Hit:vpmTimer.PulseSw 42:swp42.TransX = -5:sw42Step = 0:Me.TimerEnabled = 1:Me.TimerInterval = 10:PlaySound "fx_target",0,1,0.1,0.25:End Sub
Sub sw42_Timer()
	Select Case sw42Step
		Case 1:swp42.TransX = 3
        Case 2:swp42.TransX = -2
        Case 3:swp42.TransX = 1
        Case 4:swp42.TransX = 0:Me.TimerEnabled = 0
     End Select
	sw42Step = sw42Step + 1
End Sub

Sub sw43_Hit:vpmTimer.PulseSw 43:swp43.TransX = -5:sw43Step = 0:Me.TimerEnabled = 1:me.timerinterval = 10:PlaySound "fx_target", 0, 1, 0, 0.25:End Sub
Sub sw43_Timer()
	Select Case sw43Step
		Case 1:swp43.TransX = 3
        Case 2:swp43.TransX = -2
        Case 3:swp43.TransX = 1
        Case 4:swp43.TransX = 0:Me.TimerEnabled = 0
     End Select
	sw43Step = sw43Step + 1
End Sub

Sub sw44_Hit:vpmTimer.PulseSw 44:swp44.TransX = -5:sw44Step = 0:Me.TimerEnabled = 1:me.timerinterval = 10:PlaySound "fx_target", 0, 1, 0, 0.25:End Sub
Sub sw44_Timer()
	Select Case sw44Step
		Case 1:swp44.TransX = 3
        Case 2:swp44.TransX = -2
        Case 3:swp44.TransX = 1
        Case 4:swp44.TransX = 0:Me.TimerEnabled = 0
     End Select
	sw44Step = sw44Step + 1
End Sub

Sub sw75_Hit:vpmTimer.PulseSw 75:swp75.TransX = -5:sw75Step = 0:Me.TimerEnabled = 1:Me.TimerInterval = 10:PlaySound "fx_target", 0, 1, 0, 0.25:If Mars_Mod = 1 Then can2.visible = 1:UFO_Mars.enabled = 1:End If:End Sub
Sub sw75_Timer()
	Select Case sw75Step
		Case 1:swp75.TransX = 3
        Case 2:swp75.TransX = -2
        Case 3:swp75.TransX = 1
        Case 4:swp75.TransX = 0:Me.TimerEnabled = 0
     End Select
	sw75Step = sw75Step + 1
End Sub

Sub sw76_Hit:vpmTimer.PulseSw 76:swp76.TransX = -5:sw76Step = 0:Me.TimerEnabled = 1:Me.TimerInterval = 10:PlaySound "fx_target", 0, 1, 0, 0.25:If Mars_Mod = 1 Then can2.visible = 1:UFO_Mars.enabled = 1:End If:End Sub
Sub sw76_Timer()
	Select Case sw45Step
		Case 1:swp76.TransX = 3
        Case 2:swp76.TransX = -2
        Case 3:swp76.TransX = 1
        Case 4:swp76.TransX = 0:Me.TimerEnabled = 0
     End Select
	sw76Step = sw76Step + 1
End Sub



Sub sw45_Hit:vpmTimer.PulseSw 45:swp45.TransX = -5:sw45Step = 1:Me.TimerEnabled = 1:Me.TimerInterval = 10:PlaySound "fx_target", 0, 1, 0, 0.25:End Sub
Sub sw45_Timer()
	Select Case sw45Step
		Case 1:swp45.TransX = 3
        Case 2:swp45.TransX = -2
        Case 3:swp45.TransX = 1
        Case 4:swp45.TransX = 0:Me.TimerEnabled = 0
     End Select
	sw45Step = sw45Step + 1
End Sub



Sub sw46_Hit:vpmTimer.PulseSw 46:swp46.TransX = -5:sw46Step = 1:Me.TimerEnabled = 1:Me.TimerInterval = 10:PlaySound "fx_target", 0, 1, 0, 0.25:End Sub
Sub sw46_Timer()
	Select Case sw46Step
		Case 1:swp46.TransX = 3
        Case 2:swp46.TransX = -2
        Case 3:swp46.TransX = 1
        Case 4:swp46.TransX = 0:Me.TimerEnabled = 0
     End Select
	sw46Step = sw46Step + 1
End Sub



Sub sw47_Hit:vpmTimer.PulseSw 47:swp47.TransX = -5:sw47Step = 1:Me.TimerEnabled = 1:Me.TimerInterval = 10:PlaySound "fx_target", 0, 1, 0, 0.25:End Sub
Sub sw47_Timer()
	Select Case sw47Step
		Case 1:swp47.TransX = 3
        Case 2:swp47.TransX = -2
        Case 3:swp47.TransX = 1
        Case 4:swp47.TransX = 0:Me.TimerEnabled = 0
     End Select
	sw47Step = sw47Step + 1
End Sub

' Droptarget

Sub sw77_Hit:dtDrop.Hit 1:LEDSpeedFast:If Mars_Mod = 1 Then can2.visible = 1:UFO_Mars.enabled = 1:End If:End Sub

' Gates
Sub RGate_Hit():PlaySound "fx_gate",0,1,0.1,0.25:End Sub
Sub LGate_Hit():PlaySound "fx_gate",0,1,-0.1,0.25:End Sub

' Ramps helpers
Sub RHelp1_Hit()
    ActiveBall.VelZ = -2
    ActiveBall.VelY = 0
    ActiveBall.VelX = 0
    StopSound "fx_metalrolling"
    PlaySound "ball_bounce"
End Sub

Sub RHelp2_Hit()
    ActiveBall.VelZ = -2
    ActiveBall.VelY = 0
    ActiveBall.VelX = 0
    StopSound "fx_metalrolling"
    PlaySound "ball_bounce"
End Sub

'*********
'Solenoids
'*********

SolCallback(1) = "Auto_Plunger"
SolCallback(2) = "SolRelease"
SolCallback(3) = "bsL.SolOut"
SolCallback(4) = "bsR.SolOut"
SolCallback(5) = "SolAlien5"
SolCallback(6) = "SolAlien6"
SolCallback(7) = "vpmSolSound ""fx_Knocker"","
SolCallback(8) = "SolAlien8"
SolCallback(14) = "SolAlien14"
SolCallBack(15) = "SolUfoShake"
SolCallback(16) = "dtDrop.SolDropUp"
SolCallback(17) = "Sol17"
SolCallback(18) = "Sol18"
SolCallback(19) = "Sol19"
SolCallBack(20) = "Sol20"
SolCallback(21) = "Sol21"
SolCallback(22) = "Sol22"
SolCallback(23) = "Sol23" ' SolUfoFlash"
'SolCallback(24) = "SolBank" 'used in the Mech
SolCallback(25) = "Sol25"
SolCallback(26) = "Sol26"
SolCallback(27) = "Sol27"
SolCallBack(28) = "Sol28"
SolCallback(33) = "vpmSolGate RGate,false,"
SolCallback(34) = "vpmSolGate LGate,false,"
SolCallback(36) = "SolDiv"
SolCallback(43) = "SolStrobe"

Sub Sol17(enabled)
If enabled Then
FL17.state = 1
FL17a.state = 1
FL17b.state = 1
If BlackLightFlashers = 1 Then
ufo6b1.visible = 1
ufo6b2.visible = 1
End If
If BLWALL = 1 Then
BWONR.isdropped = 0
Outbox.image = "afm_back_right_fade"
End If
else
FL17.state = 0
FL17a.state = 0
Fl17b.state = 0
BWONR.isdropped = 1
ufo6b1.visible = 0
ufo6b2.visible = 0
Outbox.image = "afm_all_off_dark"
End If
End Sub

Sub Sol18(enabled)
If enabled Then
FL18.state = 1
FL18a.state = 1
FL18b.state = 1
If BlackLightFlashers = 1 Then
ufo5b1.visible = 1
ufo5b2.visible = 1
End If
If UfoAmbient = 1 Then
FL18UFO.state = 1
End If
If BLWALL = 1 Then
BWONR.isdropped = 0
Outbox.image = "afm_mid_right_fade"
End If
else
Outbox.image = "afm_all_off_dark"
ufo5b1.visible = 0
ufo5b2.visible = 0
BWONR.isdropped = 1
FL18.state = 0
FL18a.state = 0
Fl18b.state = 0
FL18UFO.state = 0
End If
End Sub

Sub Sol19(enabled)
If enabled Then
FL19.state = 1
FL19a.state = 1
FL19b.state = 1
If BlackLightFlashers = 1 Then
ufo4b1.visible = 1
ufo4b2.visible = 1
End If
If BLWALL = 1 Then
BWON.isdropped = 0
Outbox.image = "afm_front_right_fade"
End If
If InnerCabinetFlash = 1 AND BlackLightFlashers = 1 Then
'OutBox.image = "aBlue"
else
If InnerCabinetFlash = 1 AND BlackLightFlashers = 0 Then
'OutBox.image = "aRed"
End If
End If
If WireFormFlash = 1 AND BlackLightFlashers = 1 Then
RW.material = "BLUE"
else
If WireFormFlash = 1 AND BlackLightFlashers = 0 Then
RW.material = "RED"
End If
End If
else
Outbox.image = "afm_all_off_dark"
BWON.isdropped = 1
ufo4b1.visible = 0
ufo4b2.visible = 0
FL19.state = 0
FL19a.state = 0
Fl19b.state = 0
'OutBox.image = "Black Wood l"
RW.material = "Metal0.2"
End If
End Sub

Sub Sol20(enabled)
If enabled Then
FL20.state = 1
FL20a.state = 1
FL20ab.state = 1
FL20ac.state = 1
FL20ad.state = 1
If InnerCabinetFlash = 1 Then
'OutBox.image = "aGreen"
End If
If WireFormFlash = 1 Then
RW.material = "GREEN"
End If
For each xx in MLights:xx.Intensity = 1:next
else
FL20.state = 0
FL20a.state = 0
FL20ab.state = 0
FL20ac.state = 0
FL20ad.state = 0
'OutBox.image = "Black Wood l"
RW.material = "Metal0.2"
For each xx in MLights:xx.Intensity = 5:next
End If
End Sub

Sub Sol21(enabled)
If enabled Then
FL21.state = 1
FL21a.state = 1
else
FL21.state = 0
FL21a.state = 0
End If
End Sub

Sub Sol22(enabled)
If enabled Then
FL22.state = 1
FL22a.state = 1
else
FL22.state = 0
FL22a.state = 0
End If
End Sub

Sub Sol23(enabled)
If enabled Then
FL23.state = 1
ufo1d.image = "cupulaverde_b"
'FL23a.state = 1
If desktopmode Then
Can1.visible = 1
End If
If InnerCabinetFlash = 1 Then
'OutBox.image = "aGreen"
End If
else
FL23.state = 0
'FL23a.state = 0
Can1.visible = 0
ufo1d.image = "cupulaverde"
'OutBox.image = "Black Wood l"
End If
End Sub

Sub Sol27(enabled)
If enabled Then
FL27.state = 1
FL27a.state = 1
FL27b.state = 1
If BLWALL = 1 Then
BWON.isdropped = 0
Outbox.image = "afm_mid_left_fade"
End If
If BlackLightFlashers = 1 Then
ufo2b1.visible = 1
ufo2b2.visible = 1
End If
If InnerCabinetFlash = 1 AND BlackLightFlashers = 1 Then
'OutBox.image = "aBlue"
else
If InnerCabinetFlash = 1 AND BlackLightFlashers = 0 Then
'OutBox.image = "aRed"
End If
End If
If UfoAmbient = 1 Then
FL27UFO.state = 1
End If
If WireFormFlash = 1 AND BlackLightFlashers = 1 Then
LW.material = "BLUE"
else
If WireFormFlash = 1 AND BlackLightFlashers = 0 Then
LW.material = "RED"
End If
End If
else
FL27.state = 0
FL27a.state = 0
FL27b.state = 0
BWON.isdropped = 1
ufo2b1.visible = 0
ufo2b2.visible = 0
Outbox.image = "afm_all_off_dark"
'OutBox.image = "Black Wood l"
FL27UFO.state = 0
LW.material = "Metal0.2"
End If
End Sub

Sub Sol25(enabled)
If enabled Then
FL25.state = 1
FL25a.state = 1
FL25b.state = 1
If BLWALL = 1 Then
BWONL.isdropped = 0
Outbox.image = "afm_back_left_fade"
End If
If BlackLightFlashers = 1 Then
ufo7b1.visible = 1
ufo7b2.visible = 1
End If
else
BWONL.isdropped = 1
Outbox.image = "afm_all_off_dark"
FL25.state = 0
FL25a.state = 0
FL25b.state = 0
ufo7b1.visible = 0
ufo7b2.visible = 0
End If
End Sub

Sub Sol26(enabled)
If enabled Then
FL26.state = 1
FL26a.state = 1
FL26b.state = 1
If BlackLightFlashers = 1 Then
ufo8b1.visible = 1
ufo8b2.visible = 1
End If
If BLWALL = 1 Then
BWONM.isdropped = 0
End If
If UfoAmbient = 1 Then
FL26UFO.state = 1
End If
else
BWONM.isdropped = 1
ufo8b1.visible = 0
ufo8b2.visible = 0
FL26.state = 0
FL26a.state = 0
FL26b.state = 0
FL26UFO.state = 0
End If
End Sub

Sub Sol28(enabled)
If enabled Then
FL28.state = 1
FL28a.state = 1
FL28b.state = 1
FL28c.state = 1
If InnerCabinetFlash = 1 Then
'OutBox.image = "aGreen"
End If
If WireFormFlash = 1 Then
LW.material = "GREEN"
End If
For each xx in MLights:xx.Intensity = 1:next
else
FL28.state = 0
FL28a.state = 0
FL28a.state = 0
FL28b.state = 0
'OutBox.image = "Black Wood l"
LW.material = "Metal0.2"
For each xx in MLights:xx.Intensity = 5:next
End If
End Sub

Sub SolStrobe(enabled)
If enabled Then
FL30.state = 1
FL30a.state = 1
If BLWALL = 1 Then
'If InnerCabinetFlash = 1 Then
'OutBox.image = "aGray"
Outbox.image = "afm_all_on_dark"
End If
else
FL30.state = 0
FL30a.state = 0
'OutBox.image = "Black Wood l"
Outbox.image = "afm_all_off_dark"
End If
End Sub

Sub SolRelease(Enabled)
    If Enabled And bsTrough.Balls> 0 Then
        vpmTimer.PulseSw 31
        bsTrough.ExitSol_On
    End If
End Sub

Sub Auto_Plunger(Enabled)
    If Enabled Then
        PlungerIM.AutoFire
    End If
End Sub

Sub SolAlien5(Enabled)
    PlaySound "fx_rubber",0,1,0,0.25
    If Enabled Then
    afstep = 1
    alienf.enabled = 1      
    Else
    afstep = 2
    alienf.enabled = 1      
    End If
End Sub

Dim afstep

Sub alienf_Timer()
Select Case afstep
Case 1:
If Alien5.Transy => 20 Then
me.enabled = 0
End If
Alien5.Transy = Alien5.Transy + 1
Case 2:
If Alien5.Transy <= 0 Then
me.enabled = 0
End If
Alien5.Transy = Alien5.Transy - 1
End Select
End Sub

Sub SolAlien6(Enabled)
    PlaySound "fx_rubber",0,1,0,0.25
    If Enabled Then
    asstep = 1
    aliens.enabled = 1      
    Else
    asstep = 2
    aliens.enabled = 1      
    End If
End Sub

Dim asstep

Sub aliens_Timer()
Select Case asstep
Case 1:
If Alien6.Transy => 20 Then
me.enabled = 0
End If
Alien6.Transy = Alien6.Transy + 1
Case 2:
If Alien6.Transy <= 0 Then
me.enabled = 0
End If
Alien6.Transy = Alien6.Transy - 1
End Select
End Sub

Sub SolAlien8(Enabled)
    PlaySound "fx_rubber",0,1,0,0.25
    If Enabled Then
    aestep = 1
    aliene.enabled = 1      
    Else
    aestep = 2
    aliene.enabled = 1      
    End If
End Sub

Dim aestep

Sub aliene_Timer()
Select Case aestep
Case 1:
If Alien8.Transy => 20 Then
me.enabled = 0
End If
Alien8.Transy = Alien8.Transy + 1
Case 2:
If Alien8.Transy <= 0 Then
me.enabled = 0
End If
Alien8.Transy = Alien8.Transy - 1
End Select
End Sub

Sub SolAlien14(Enabled)
    PlaySound "fx_rubber",0,1,0,0.25
    If Enabled Then
    atfstep = 1
    alientf.enabled = 1      
    Else
    atfstep = 2
    alientf.enabled = 1      
    End If
End Sub

Dim atfstep

Sub alientf_Timer()
Select Case atfstep
Case 1:
If Alien14.Transy => 20 Then
me.enabled = 0
End If
Alien14.Transy = Alien14.Transy + 1
Case 2:
If Alien14.Transy <= 0 Then
me.enabled = 0
End If
Alien14.Transy = Alien14.Transy - 1
End Select
End Sub

'************************
' Diverter animation
'************************

Sub SolDiv(Enabled)
    If Enabled Then
        DiverterDir = 1
    Else
        DiverterDir = -1
    End If

    Diverter.Enabled = 0
    If DiverterPos <1 Then DiverterPos = 1
    If DiverterPos> 5 Then DiverterPos = 5

    Diverter.Enabled = 1
End Sub

Sub Diverter_Timer()
    Select Case DiverterPos
        Case 0:Diverter1.IsDropped = 0:Diverter2.IsDropped = 1:Diverter.Enabled = 0
        Case 1:Diverter2.IsDropped = 0:Diverter1.IsDropped = 1:Diverter3.IsDropped = 1
        Case 2:Diverter3.IsDropped = 0:Diverter2.IsDropped = 1:Diverter4.IsDropped = 1
        Case 3:Diverter4.IsDropped = 0:Diverter3.IsDropped = 1:Diverter5.IsDropped = 1
        Case 4:Diverter5.IsDropped = 0:Diverter4.IsDropped = 1:Diverter6.IsDropped = 1
        Case 5:Diverter6.IsDropped = 0:Diverter5.IsDropped = 1
        Case 6:Diverter.Enabled = 0
    End Select

    DiverterPos = DiverterPos + DiverterDir
End Sub

'********************
' Special JP Flippers
'********************

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
    Dim tmp, tmp2
    If Enabled Then
        PlaySound "fx_flipperup", 0, 1, -0.1, 0.25
        LeftFlipper.RotateToEnd
    Else
'        tmp = LeftFlipper.Strength
'        tmp2 = LeftFlipper.Recoil
'        LeftFlipper.Strength = 6 'increase return strength to compensate for the slower speed
'        LeftFlipper.Recoil = 0.5 'decrease the recoil to allow drop catches
        PlaySound "fx_flipperdown", 0, 1, -0.1, 0.25
        LeftFlipper.RotateToStart
'        LeftFlipper.Strength = tmp
'        LeftFlipper.Recoil = tmp2
    End If
End Sub

Sub SolRFlipper(Enabled)
    Dim tmp, tmp2
    If Enabled Then
        PlaySound "fx_flipperup", 0, 1, 0.1, 0.25
        RightFlipper.RotateToEnd
    Else
'        tmp = RightFlipper.Strength
'        tmp2 = LeftFlipper.Recoil
'        RightFlipper.Strength = 6 'increase return strength to compensate for the slower speed
'        RightFlipper.Recoil = 0.5 'decrease the recoil to allow drop catches
        PlaySound "fx_flipperdown", 0, 1, 0.1, 0.25
        RightFlipper.RotateToStart
'        RightFlipper.Strength = tmp
'        RightFlipper.Recoil = tmp2
    End If
End Sub

Sub LeftFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 10, -0.1, 0.25
End Sub

Sub RightFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 10, 0.1, 0.25
End Sub

'***********************
'   lipper Primitives
'***********************

Sub Timer_Timer()
    'prgate.Rotx = rgater.CurrentAngle
    'pLgate.Rotx = RgateL.CurrentAngle

    flipperl.RotY = LeftFlipper.CurrentAngle
    flipperr.RotY = RightFlipper.CurrentAngle
End Sub

'******************
'Motor Bank Up Down
'******************

Sub Update3Bank(currpos, currspeed, lastpos)
    Dim zpos
    zpos = 0 - currpos
    If currpos <> lastpos Then
        BackBank.Z = zpos - 10
        swp45.Z = zpos - 10
        swp46.Z = zpos - 10
        swp47.Z = zpos - 10
    End If

    If currpos = 50 Then
        sw45.Isdropped = 1
        sw46.Isdropped = 1
        sw47.Isdropped = 1
        LEDSpeedMedium
    End If
    If currpos = 0 Then
        sw45.Isdropped = 0
        sw46.Isdropped = 0
        sw47.Isdropped = 0
        LEDSpeedSlow
    End If
End Sub

'***********
' Update GI
'***********

Dim gistep,xx,s1,s2,s3,s4,s5,s6,s8,t1,t2,t3,t4,t5,t6,t8
gistep = 1/8

Sub UpdateGI(no, step)

If step = 0 then exit sub

    Select Case no

        Case 0

For each xx in BGI:xx.IntensityScale = gistep * step:next

If step > 4 Then
For each xx in LLights:xx.Intensity = 1:next
End If

If step < 4 Then
For each xx in LLights:xx.Intensity = 10:next
End If

If cController = 3 Then

If step = 1 and s1 = 0 then Controller.B2SSetData 101,0:Controller.B2SSetData 102,0:Controller.B2SSetData 103,0:Controller.B2SSetData 104,0:Controller.B2SSetData 105,0:Controller.B2SSetData 106,0:s1=1:s2 = 0:s3 = 0:s4 = 0:s5 = 0:s6 = 0:s8 = 0
If step = 2 and s2 = 0 then Controller.B2SSetData 101,1:Controller.B2SSetData 102,0:Controller.B2SSetData 103,0:Controller.B2SSetData 104,0:Controller.B2SSetData 105,0:Controller.B2SSetData 106,0:s1=0:s2 = 1:s3 = 0:s4 = 0:s5 = 0:s6 = 0:s8 = 0
If step = 3 and s3 = 0 then Controller.B2SSetData 102,1:Controller.B2SSetData 101,0:Controller.B2SSetData 103,0:Controller.B2SSetData 104,0:Controller.B2SSetData 105,0:Controller.B2SSetData 106,0:s1=0:s2 = 0:s3 = 1:s4 = 0:s5 = 0:s6 = 0:s8 = 0
If step = 4 and s4 = 0 then Controller.B2SSetData 103,1:Controller.B2SSetData 102,0:Controller.B2SSetData 101,0:Controller.B2SSetData 104,0:Controller.B2SSetData 105,0:Controller.B2SSetData 106,0:s1=0:s2 = 0:s3 = 0:s4 = 1:s5 = 0:s6 = 0:s8 = 0
If step = 5 and s5 = 0 then Controller.B2SSetData 104,1:Controller.B2SSetData 102,0:Controller.B2SSetData 103,0:Controller.B2SSetData 101,0:Controller.B2SSetData 105,0:Controller.B2SSetData 106,0:s1=0:s2 = 0:s3 = 0:s4 = 0:s5 = 1:s6 = 0:s8 = 0
If step = 6 and s6 = 0 then Controller.B2SSetData 105,1:Controller.B2SSetData 102,0:Controller.B2SSetData 103,0:Controller.B2SSetData 104,0:Controller.B2SSetData 101,0:Controller.B2SSetData 106,0:s1=0:s2 = 0:s3 = 0:s4 = 0:s5 = 0:s6 = 1:s8 = 0
If step = 8 and s8 = 0 then Controller.B2SSetData 106,1:Controller.B2SSetData 102,0:Controller.B2SSetData 103,0:Controller.B2SSetData 104,0:Controller.B2SSetData 105,0:Controller.B2SSetData 101,0:s1=0:s2 = 0:s3 = 0:s4 = 0:s5 = 0:s6 = 0:s8 = 1

End If

        Case 1

For each xx in MGI:xx.IntensityScale = gistep * step:next

If step > 4 Then
For each xx in MLights:xx.Intensity = 1:next
End If

If step < 4 Then
For each xx in MLights:xx.Intensity = 10:next
End If

If cController = 3 Then

If step = 1 and t1 = 0 then Controller.B2SSetData 111,0:Controller.B2SSetData 112,0:Controller.B2SSetData 113,0:Controller.B2SSetData 114,0:Controller.B2SSetData 115,0:Controller.B2SSetData 116,0:t1=1:t2 = 0:t3 = 0:t4 = 0:t5 = 0:t6 = 0:t8 = 0
If step = 2 and t2 = 0 then Controller.B2SSetData 111,1:Controller.B2SSetData 112,0:Controller.B2SSetData 113,0:Controller.B2SSetData 114,0:Controller.B2SSetData 115,0:Controller.B2SSetData 116,0:t1=0:t2 = 1:t3 = 0:t4 = 0:t5 = 0:t6 = 0:t8 = 0
If step = 3 and t3 = 0 then Controller.B2SSetData 112,1:Controller.B2SSetData 111,0:Controller.B2SSetData 113,0:Controller.B2SSetData 114,0:Controller.B2SSetData 115,0:Controller.B2SSetData 116,0:t1=0:t2 = 0:t3 = 1:t4 = 0:t5 = 0:t6 = 0:t8 = 0
If step = 4 and t4 = 0 then Controller.B2SSetData 113,1:Controller.B2SSetData 112,0:Controller.B2SSetData 111,0:Controller.B2SSetData 114,0:Controller.B2SSetData 115,0:Controller.B2SSetData 116,0:t1=0:t2 = 0:t3 = 0:t4 = 1:t5 = 0:t6 = 0:t8 = 0
If step = 5 and t5 = 0 then Controller.B2SSetData 114,1:Controller.B2SSetData 112,0:Controller.B2SSetData 113,0:Controller.B2SSetData 111,0:Controller.B2SSetData 115,0:Controller.B2SSetData 116,0:t1=0:t2 = 0:t3 = 0:t4 = 0:t5 = 1:t6 = 0:t8 = 0
If step = 6 and t6 = 0 then Controller.B2SSetData 115,1:Controller.B2SSetData 112,0:Controller.B2SSetData 113,0:Controller.B2SSetData 114,0:Controller.B2SSetData 111,0:Controller.B2SSetData 116,0:t1=0:t2 = 0:t3 = 0:t4 = 0:t5 = 0:t6 = 1:t8 = 0
If step = 8 and t8 = 0 then Controller.B2SSetData 116,1:Controller.B2SSetData 112,0:Controller.B2SSetData 113,0:Controller.B2SSetData 114,0:Controller.B2SSetData 115,0:Controller.B2SSetData 111,0:t1=0:t2 = 0:t3 = 0:t4 = 0:t5 = 0:t6 = 0:t8 = 1

End If

       Case 2

For each xx in TGI:xx.IntensityScale = gistep * step:next

If step > 4 Then
For each xx in TLights:xx.Intensity = 1:next
End If

If step < 4 Then
For each xx in TLights:xx.Intensity = 10:next
End If

    End Select
End Sub

'******************
' UFO Shake & Leds
'******************
Dim UfoLedPos, cBall, ufoalternate
BigUfoInit

ufoalternate = 1


Sub BigUfoInit
	'UfoLedPos = 0
    Set cBall = ckicker.createball
    ckicker.Kick 0, 0
End Sub

Sub UFOLed_Timer()
	If BlackLights = 1 then 
   Select Case UfoLedPos
			Case 0:ufo1.image = "bigufo3_bl":UfoLedPos = 1
			Case 1:ufo1.image = "bigufo2_bl":UfoLedPos = 2
			Case 2:ufo1.image = "bigufo1_bl":UfoLedPos = 0
    End Select
		Else
   Select Case UfoLedPos
			Case 0:ufo1.image = "bigufo3":UfoLedPos = 1
			Case 1:ufo1.image = "bigufo2":UfoLedPos = 2
			Case 2:ufo1.image = "bigufo1":UfoLedPos = 0
    End Select
		End If

'ufo1.TriggerSingleUpdate
End Sub

Sub SolUfoShake(Enabled)
    If Enabled Then
        BigUfoShake
    End If
End Sub

Sub BigUfoShake
    cball.velx = 10 + 2*RND(1)
    cball.vely = 2*(RND(1)-RND(1))
End Sub

Sub SmallUfoShake
    cball.velx = 2*RND(1)
    cball.vely = 2*(RND(1)-RND(1))
End Sub

Sub BigUfoUpdate
    Ufo1.rotx = -(ckicker.y - cball.y)
    Ufo1.transx = (ckicker.y - cball.y)/2
    Ufo1.roty = cball.x - ckicker.x
    
	Ufo3.rotx = -(ckicker.y - cball.y)
    Ufo3.transx = (ckicker.y - cball.y)/2
    Ufo3.roty = cball.x - ckicker.x

	Ufo9.rotx = -(ckicker.y - cball.y)
    Ufo9.transx = (ckicker.y - cball.y)/2
    Ufo9.roty = cball.x - ckicker.x

	Ufo10.rotx = -(ckicker.y - cball.y)
    Ufo10.transx = (ckicker.y - cball.y)/2
    Ufo10.roty = cball.x - ckicker.x

	Ufo11.rotx = -(ckicker.y - cball.y)
    Ufo11.transx = (ckicker.y - cball.y)/2
    Ufo11.roty = cball.x - ckicker.x

    Ufo12.rotx = -(ckicker.y - cball.y)
    Ufo12.transx = (ckicker.y - cball.y)/2
    Ufo12.roty = cball.x - ckicker.x

	Ufo13.rotx = -(ckicker.y - cball.y)
    Ufo13.transx = (ckicker.y - cball.y)/2
    Ufo13.roty = cball.x - ckicker.x

	Ufo14.rotx = -(ckicker.y - cball.y)
    Ufo14.transx = (ckicker.y - cball.y)/2
    Ufo14.roty = cball.x - ckicker.x

	Ufo15.rotx = -(ckicker.y - cball.y)
    Ufo15.transx = (ckicker.y - cball.y)/2
    Ufo15.roty = cball.x - ckicker.x

	Ufo16.rotx = -(ckicker.y - cball.y)
    Ufo16.transx = (ckicker.y - cball.y)/2
    Ufo16.roty = cball.x - ckicker.x

    can1.rotx = -(ckicker.y - cball.y)
    can1.transx = (ckicker.y - cball.y)/2
    can1.roty = cball.x - ckicker.x

    can2.rotx = -(ckicker.y - cball.y)
    can2.transx = (ckicker.y - cball.y)/2
    can2.roty = cball.x - ckicker.x

    Ufo1d.rotx = (ckicker.y - cball.y)
    Ufo1d.transx = -(ckicker.y - cball.y)/2
    Ufo1d.roty = cball.x - ckicker.x
End Sub

'******************
' RealTime Updates
'******************
Set MotorCallback = GetRef("GameTimer")

Sub GameTimer
    BigUfoUpdate
    'RollingSound
    'UpdateFlipperLogos
End Sub

Sub Saucer_LEDS_Timer()

If L101.State = 1 Then
ufo1.visible = 1
ufo15.visible = 1
else
ufo1.visible = 0
ufo15.visible = 0
End If

If L102.State = 1 Then
ufo3.visible = 1
else
ufo3.visible = 0
End If

If L103.State = 1 Then
ufo9.visible = 1
else
ufo9.visible = 0
End If

If L104.State = 1 Then
ufo10.visible = 1
else
ufo10.visible = 0
End If

If L105.State = 1 Then
ufo11.visible = 1
else
ufo11.visible = 0
End If

If L106.State = 1 Then
ufo12.visible = 1
ufo16.visible = 1
else
ufo12.visible = 0
ufo16.visible = 0
End If

If L107.State = 1 Then
ufo13.visible = 1
else
ufo13.visible = 0
End If

If L101.State + L102.State + L103.State + L104.State + L105.State + L106.State + L107.State = 0 Then
else
End If
End Sub

'******************************
' Small and Big UFOs LEDs Speed
'******************************

Sub LEDSpeedSlow()
    Dim Obj
    UfoLed.Interval = 800
    UFOLightTimer.Interval = 100
End Sub

Sub LEDSpeedMedium()
    Dim obj
    UfoLed.Interval = 500
    UFOLightTimer.Interval = 50
End Sub

Sub LEDSPeedFast()
    Dim obj
    UfoLed.Interval = 300
    UFOLightTimer.Interval = 25
End Sub

'****************
' Small UFOS Leds
'****************

Dim UFOSmallPos
UFOSmallPos = 0

Sub UFOLightTimer_Timer()
if ufoalternate then

If SpinningSaucers = 1 then
ufo2.RotZ = ufo2.RotZ +2
ufo4.RotZ = ufo2.RotZ +2
ufo5.RotZ = ufo2.RotZ +2
ufo6.RotZ = ufo2.RotZ +2
ufo7.RotZ = ufo2.RotZ +2
ufo8.RotZ = ufo2.RotZ +2
ufo2b.RotZ = ufo2b.RotZ -2
ufo2b1.RotZ = ufo2b1.RotZ -2
ufo2b2.RotZ = ufo2b2.RotZ -2
ufo4b.RotZ = ufo2b.RotZ -2
ufo4b1.RotZ = ufo2b.RotZ -2
ufo5b.RotZ = ufo2b.RotZ -2
ufo5b1.RotZ = ufo2b.RotZ -2
ufo6b.RotZ = ufo2b.RotZ -2
ufo6b1.RotZ = ufo2b.RotZ -2
ufo7b.RotZ = ufo2b.RotZ -2
ufo7b1.RotZ = ufo2b.RotZ -2
ufo7b2.RotZ = ufo2b.RotZ -2
ufo8b.RotZ = ufo2b.RotZ -2
ufo8b1.RotZ = ufo2b.RotZ -2

End If
else
    Select Case UFOSmallPos
        Case 0:ufo2.imageA = "ufo-center1":ufo4.imageA = "ufo-center4":ufo5.imageA = "ufo-center5":ufo6.imageA = "ufo-center6":ufo7.imageA = "ufo-center7":ufo8.imageA = "ufo-center8":UFOSmallPos = 1
        Case 1:ufo2.imageA = "ufo-center2":ufo4.imageA = "ufo-center5":ufo5.imageA = "ufo-center6":ufo6.imageA = "ufo-center7":ufo7.imageA = "ufo-center8":ufo8.imageA = "ufo-center1":UFOSmallPos = 2
        Case 2:ufo2.imageA = "ufo-center3":ufo4.imageA = "ufo-center6":ufo5.imageA = "ufo-center7":ufo6.imageA = "ufo-center8":ufo7.imageA = "ufo-center1":ufo8.imageA = "ufo-center2":UFOSmallPos = 3
        Case 3:ufo2.imageA = "ufo-center4":ufo4.imageA = "ufo-center7":ufo5.imageA = "ufo-center8":ufo6.imageA = "ufo-center1":ufo7.imageA = "ufo-center2":ufo8.imageA = "ufo-center3":UFOSmallPos = 4
        Case 4:ufo2.imageA = "ufo-center5":ufo4.imageA = "ufo-center8":ufo5.imageA = "ufo-center1":ufo6.imageA = "ufo-center2":ufo7.imageA = "ufo-center3":ufo8.imageA = "ufo-center4":UFOSmallPos = 5
        Case 5:ufo2.imageA = "ufo-center6":ufo4.imageA = "ufo-center1":ufo5.imageA = "ufo-center2":ufo6.imageA = "ufo-center3":ufo7.imageA = "ufo-center4":ufo8.imageA = "ufo-center5":UFOSmallPos = 6
        Case 6:ufo2.imageA = "ufo-center7":ufo4.imageA = "ufo-center2":ufo5.imageA = "ufo-center3":ufo6.imageA = "ufo-center4":ufo7.imageA = "ufo-center5":ufo8.imageA = "ufo-center6":UFOSmallPos = 7
        Case 7:ufo2.imageA = "ufo-center8":ufo4.imageA = "ufo-center3":ufo5.imageA = "ufo-center4":ufo6.imageA = "ufo-center5":ufo7.imageA = "ufo-center6":ufo8.imageA = "ufo-center7":UFOSmallPos = 0
    End Select
End if
End Sub

''''''''''''''''''''''''''''''''''''''''''
''''''''Blacklights'''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''

Sub SetBlacklights ()
	If BlackLights = 1 then
		'ufo1.image = "bigufo1_bl"
		ufo2b.image = "ufo-small_bl_on"
		ufo4b.image = "ufo-small_bl"
		ufo5b.image = "ufo-small_bl"
		ufo6b.image = "ufo-small_bl"
		ufo7b.image = "ufo-small_bl"
		ufo8b.image = "ufo-small_bl"
	Else
		'ufo1.image = "bigufo1"
		ufo2b.image = "ufo-small"
		ufo4b.image = "ufo-small"
		ufo5b.image = "ufo-small"
		ufo6b.image = "ufo-small"
		ufo7b.image = "ufo-small"
		ufo8b.image = "ufo-small"
	End If
End Sub

'***************************************************
'  JP's Fading Lamps & Flashers version 9 for VP921
'   Based on PD's Fading Lights
' SetLamp 0 is Off
' SetLamp 1 is On
' FadingLevel(x) = fading state
' LampState(x) = light state
' Includes the flash element (needs own timer)
' Flashers can be used as lights too
'***************************************************

Dim LampState(200), FadingLevel(200)
Dim FlashState(200), FlashLevel(200)
Dim FlashSpeedUp, FlashSpeedDown
Dim FadeArray1: FadeArray1 = Array("BlackHole_on", "BlackHole_66", "BlackHole_33", "BlackHole_off")
Dim YellowLight: YellowLight = Array("yellow_on", "yellow_66", "yellow_33", "yellow_off")
Dim OrangeLight: orangeLight = Array("orange_on", "orange_66", "orange_33", "orange_off")
Dim RedLight: RedLight = Array("Red_on", "Red_66", "Red_33", "Red_off")
Dim BlueLightL: BlueLightL = Array("bluearrowL_on", "bluearrowL_66", "bluearrowL_33", "bluearrowL_off")
Dim BlueLightR: BlueLightR = Array("bluearrowR_on", "bluearrowR_66", "bluearrowR_33", "bluearrowR_off")
Dim BArray: BArray = Array("BH_Bumpercap_on", "BH_Bumpercap_66", "BH_Bumpercap_33", "BH_Bumpercap_off")
Dim BArrayDay: BArrayDay = Array("BH_Bumpercap_Day_on", "BH_Bumpercap_Day_66", "BH_Bumpercap_Day_33", "BH_Bumpercap_Day_off")
Dim ReEnetryTubeArray: ReEnetryTubeArray = Array("BH_ReEnetryTube_on", "BH_ReEnetryTube_66", "BH_ReEnetryTube_33", "BH_ReEnetryTube_off")

Const LightHaloBrightness		= 100


FlashInit()
FlasherTimer.Interval = 10 'flash fading speed
FlasherTimer.Enabled = 0

' Lamp & Flasher Timers

Sub StartLampTimer
	AllLampsOff()
	LampTimer.Interval = 30 'lamp fading speed
	LampTimer.Enabled = 1
End Sub

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
 
 Sub UpdateLamps
	NFadeL 11, l11
	NFadeL 12, l12
	NFadeL 13, l13
	NFadeL 14, l14
	NFadeL 15, l15
	NFadeL 16, l16
	NFadeL 17, l17
	NFadeL 18, l18
	NFadeL 21, l21
	NFadeL 22, l22
	NFadeL 23, l23
	NFadeL 24, l24
	NFadeL 25, l25
	NFadeL 26, l26
	NFadeL 27, l27
	NFadeL 28, l28
	NFadeL 31, l31
	NFadeL 32, l32
	NFadeL 33, l33
	NFadeL 34, l34
	NFadeL 35, l35
	NFadeL 36, l36
	NFadeL 37, l37
	NFadeL 38, l38
	NFadeL 41, l41
	NFadeL 42, l42
	NFadeLm 43, l43
	NFadeL 43, l43a
    NFadeLm 44, l44
    NFadeL 44, l44a
	NFadeL 45, l45
	NFadeL 46, l46
	NFadeL 47, l47
	NFadeL 48, l48
	NFadeL 51, l51
	NFadeL 52, l52
	NFadeL 53, l53
	NFadeL 54, l54
	NFadeL 55, l55
	NFadeL 56, l56
	NFadeL 57, l57
	NFadeL 58, l58
	NFadeL 61, l61
	NFadeL 62, l62
	NFadeL 63, l63
	NFadeL 64, l64
	NFadeL 65, l65
	NFadeL 66, l66
	NFadeL 67, l67
	NFadeL 68, l68
	NFadeL 71, l71
	NFadeL 72, l72
	NFadeL 73, l73
	NFadeL 74, l74
	NFadeL 75, l75
	NFadeL 76, l76
	NFadeL 77, l77
	NFadeL 78, l78
	NFadeL 81, l81
	NFadeL 82, l82
	NFadeL 83, l83
	NFadeL 84, l84
	NFadeL 85, l85

NFadeL 101, l101
NFadeL 102, l102
NFadeL 103, l103
NFadeL 104, l104
NFadeL 105, l105
NFadeL 106, l106
NFadeL 107, l107


End Sub

'Sindbad: You can use this instead of FadeLN
' call it this way: FadeLight lampnumber, light, Array
Sub FadeLight(nr, light, group)
	Select Case FadingLevel(nr)
		Case 2:light.State = LightStateOff:light.OffImage = group(3):FadingLevel(nr) = 0
		Case 3:light.State = LightStateOn:light.OnImage = group(2):FadingLevel(nr) = 2
		Case 4:light.State = LightStateOn:light.OnImage = group(1):FadingLevel(nr) = 3
		Case 5:light.State = LightStateOn:light.OnImage = group(0):FadingLevel(nr) = 1
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

'cyberpez FadeFlash can be used to swap images on ramps or flashers

Sub FadeFlash(nr, light, group)
	Select Case FadingLevel(nr)
		Case 2:light.imageA = group(3):FadingLevel(nr) = 0
		Case 3:light.imageA = group(2):FadingLevel(nr) = 2
		Case 4:light.imageA = group(1):FadingLevel(nr) = 3
		Case 5:light.imageA = group(0):FadingLevel(nr) = 1
	End Select
End Sub

Sub FadeFlashm(nr, light, group)
	Select Case FadingLevel(nr)
		Case 2:light.imageA = group(3):
		Case 3:light.imageA = group(2):
		Case 4:light.imageA = group(1):
		Case 5:light.imageA = group(0):
	End Select
End Sub

LightHalos_Init
Sub LightHalos_Init

End Sub

Sub FlasherTimer_Timer()

End Sub

' div lamp subs

Sub AllLampsOff()
    Dim x
    For x = 0 to 200
        LampState(x) = 0
        FadingLevel(x) = 4
'        FadingLevel2(x) = 4
    Next

UpdateLamps:UpdateLamps:Updatelamps
End Sub

Sub SetLamp(nr, value)
    If value <> LampState(nr) Then
        LampState(nr) = abs(value)
        FadingLevel(nr) = abs(value) + 4
'		FadingLevel2(nr) = abs(value) + 4
    End If
End Sub

' div flasher subs

Sub FlashInit
    Dim i
    For i = 0 to 200
        FlashState(i) = 0
        FlashLevel(i) = 0
    Next

    FlashSpeedUp = 30   ' fast speed when turning on the flasher
    FlashSpeedDown = 10 ' slow speed when turning off the flasher, gives a smooth fading
    ' you could also change the default images for each flasher or leave it as in the editor
    ' for example
    ' Flasher1.Image = "fr"
    AllFlashOff()
End Sub

Sub AllFlashOff
    Dim i
    For i = 0 to 200
        FlashState(i) = 0
    Next
End Sub

Sub SetFlash(nr, stat)
    FlashState(nr) = ABS(stat)
End Sub


' Flasher objects
' Uses own faster timer

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
            If FlashLevel(nr) > 250 Then
                FlashLevel(nr) = 250
                FlashState(nr) = -2 'completely on
            End if
            Object.opacity = FlashLevel(nr)
    End Select
End Sub

Sub Flashm(nr, object) 'multiple flashers, it doesn't change the flashstate
    Select Case FlashState(nr)
        Case 0         'off
            Object.opacity = FlashLevel(nr)
        Case 1         ' on
            Object.opacity = FlashLevel(nr)
    End Select
End Sub

Sub FlashVal(nr, object, value)
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
            If FlashLevel(nr) > value Then
                FlashLevel(nr) = value
                FlashState(nr) = -2 'completely on
            End if
            Object.opacity = FlashLevel(nr)
    End Select
End Sub

Sub FlashValm(nr, object, value) 'multiple flashers, it doesn't change the flashstate
    Select Case FlashState(nr)
        Case 0 'off
            Object.opacity = FlashLevel(nr)
        Case 1 ' on
            Object.opacity = FlashLevel(nr)
    End Select
End Sub



Sub FadeLn(nr, Light, a, b, c, d)
    Select Case FadingLevel(nr)
        Case 2:Light.Offimage = d:FadingLevel(nr) = 0 'Off
        Case 3:Light.Offimage = c:FadingLevel(nr) = 2 'fading...
        Case 4:Light.Offimage = b:FadingLevel(nr) = 3 'fading...
        Case 5:Light.Offimage = a:FadingLevel(nr) = 1 'ON
    End Select
End Sub

Sub FadeLnm(nr, Light, a, b, c, d)
    Select Case FadingLevel(nr)
        Case 2:Light.Offimage = d
        Case 3:Light.Offimage = c
        Case 4:Light.Offimage = b
        Case 5:Light.Offimage = a
    End Select
End Sub

Sub LMapn(nr, Light, a, b, c, d)
    Select Case FadingLevel(nr)
        Case 2:Light.Onimage = d:FadingLevel(nr) = 0 'Off
        Case 3:Light.ONimage = c:FadingLevel(nr) = 2 'fading...
        Case 4:Light.ONimage = b:FadingLevel(nr) = 3 'fading...
        Case 5:Light.ONimage = a:FadingLevel(nr) = 1 'ON
    End Select
End Sub

Sub LMapnm(nr, Light, a, b, c, d)
    Select Case FadingLevel(nr)
        Case 2:Light.ONimage = d
        Case 3:Light.ONimage = c
        Case 4:Light.ONimage = b
        Case 5:Light.ONimage = a
    End Select
End Sub
' Walls

Sub FadeW(nr, a, b, c)
    Select Case FadingLevel(nr)
        Case 2:c.IsDropped = 1:FadingLevel(nr) = 0                 'Off
        Case 3:b.IsDropped = 1:c.IsDropped = 0:FadingLevel(nr) = 2 'fading...
        Case 4:a.IsDropped = 1:b.IsDropped = 0:FadingLevel(nr) = 3 'fading...
        Case 5:a.IsDropped = 0:FadingLevel(nr) = 1                 'ON
    End Select
End Sub

Sub FadeWm(nr, a, b, c)
    Select Case FadingLevel(nr)
        Case 2:c.IsDropped = 1
        Case 3:b.IsDropped = 1:c.IsDropped = 0
        Case 4:a.IsDropped = 1:b.IsDropped = 0
        Case 5:a.IsDropped = 0
    End Select
End Sub

Sub NFadeW(nr, a)
    Select Case FadingLevel(nr)
        Case 4:a.IsDropped = 1:FadingLevel(nr) = 0
        Case 5:a.IsDropped = 0:FadingLevel(nr) = 1
    End Select
End Sub

Sub NFadeWm(nr, a)
    Select Case FadingLevel(nr)
        Case 4:a.IsDropped = 1
        Case 5:a.IsDropped = 0
    End Select
End Sub

Sub NFadeWi(nr, a)
    Select Case FadingLevel(nr)
        Case 4:a.IsDropped = 0:FadingLevel(nr) = 0
        Case 5:a.IsDropped = 1:FadingLevel(nr) = 1
    End Select
End Sub

Sub NFadeWim(nr, a)
    Select Case FadingLevel(nr)
        Case 4:a.IsDropped = 0
        Case 5:a.IsDropped = 1
    End Select
End Sub

'Lights

Sub FadeL(nr, a, b)
    Select Case FadingLevel(nr)
        Case 2:b.state = 0:FadingLevel(nr) = 0
        Case 3:b.state = 1:FadingLevel(nr) = 2
        Case 4:a.state = 0:FadingLevel(nr) = 3
        Case 5:a.state = 1:FadingLevel(nr) = 1
    End Select
End Sub

Sub FadeLm(nr, a, b)
    Select Case FadingLevel(nr)
        Case 2:b.state = 0
        Case 3:b.state = 1
        Case 4:a.state = 0
        Case 5:a.state = 1
    End Select
End Sub



Sub LMap(nr, a, b, c) 'can be used with normal/olod style lights too
    Select Case FadingLevel(nr)
        Case 2:c.state = 0:FadingLevel(nr) = 0
        Case 3:b.state = 0:c.state = 1:FadingLevel(nr) = 2
        Case 4:a.state = 0:b.state = 1:FadingLevel(nr) = 3
        Case 5:b.state = 0:c.state = 0:a.state = 1:FadingLevel(nr) = 1
    End Select
End Sub

Sub LMapm(nr, a, b, c)
    Select Case FadingLevel(nr)
        Case 2:c.state = 0
        Case 3:b.state = 0:c.state = 1
        Case 4:a.state = 0:b.state = 1
        Case 5:b.state = 0:c.state = 0:a.state = 1
    End Select
End Sub

'Reels

Sub FadeR(nr, a)
    Select Case FadingLevel(nr)
        Case 2:a.SetValue 3:FadingLevel(nr) = 0
        Case 3:a.SetValue 2:FadingLevel(nr) = 2
        Case 4:a.SetValue 1:FadingLevel(nr) = 3
        Case 5:a.SetValue 1:FadingLevel(nr) = 1
    End Select
End Sub

Sub FadeRm(nr, a)
    Select Case FadingLevel(nr)
        Case 2:a.SetValue 3
        Case 3:a.SetValue 2
        Case 4:a.SetValue 1
        Case 5:a.SetValue 1
    End Select
End Sub

'Texts

Sub NFadeT(nr, a, b)
    Select Case FadingLevel(nr)
        Case 4:a.Text = "":FadingLevel(nr) = 0
        Case 5:a.Text = b:FadingLevel(nr) = 1
    End Select
End Sub

Sub NFadeTm(nr, a, b)
    Select Case FadingLevel(nr)
        Case 4:a.Text = ""
        Case 5:a.Text = b
    End Select
End Sub

' Flash a light, not controlled by the rom

Sub FlashL(nr, a, b)
    Select Case FadingLevel(nr)
        Case 1:b.state = 0:FadingLevel(nr) = 0
        Case 2:b.state = 1:FadingLevel(nr) = 1
        Case 3:a.state = 0:FadingLevel(nr) = 2
        Case 4:a.state = 1:FadingLevel(nr) = 3
        Case 5:b.state = 1:FadingLevel(nr) = 4
    End Select
End Sub

' Light acting as a flash. C is the light number to be restored

Sub MFadeL(nr, a, b, c)
    Select Case FadingLevel(nr)
        Case 2:b.state = 0:FadingLevel(nr) = 0:SetLamp c, FadingLevel(c)
        Case 3:b.state = 1:FadingLevel(nr) = 2
        Case 4:a.state = 0:FadingLevel(nr) = 3
        Case 5:a.state = 1:FadingLevel(nr) = 1
    End Select
End Sub

Sub MFadeLm(nr, a, b, c)
    Select Case FadingLevel(nr)
        Case 2:b.state = 0:SetLamp c, FadingLevel(c)
        Case 3:b.state = 1
        Case 4:a.state = 0
        Case 5:a.state = 1
    End Select
End Sub

'Alpha Ramps used as fading lights
'ramp is the name of the ramp
'a,b,c,d are the images used for on...off
'r is the refresh light

Sub FadeAR(nr, ramp, a, b, c, d)
    Select Case FadingLevel(nr)
        Case 2:ramp.image = d:FadingLevel(nr) = 0 'Off
        Case 3:ramp.image = c:FadingLevel(nr) = 2 'fading...
        Case 4:ramp.image = b:FadingLevel(nr) = 3 'fading...
        Case 5:ramp.image = a:FadingLevel(nr) = 1 'ON
    End Select
End Sub

Sub FadeARm(nr, ramp, a, b, c, d)
    Select Case FadingLevel(nr)
        Case 2:ramp.image = d
        Case 3:ramp.image = c
        Case 4:ramp.image = b
        Case 5:ramp.image = a
    End Select
End Sub

Sub FlashFO(nr, ramp, a, b, c)                                   'used for reflections when there is no off ramp
    Select Case FadingLevel(nr)
        Case 2:ramp.IsVisible = 0:FadingLevel(nr) = 0                'Off
        Case 3:ramp.image = c:FadingLevel(nr) = 2                'fading...
        Case 4:ramp.image = b:FadingLevel(nr) = 3                'fading...
        Case 5:ramp.image = a:ramp.IsVisible = 1:FadingLevel(nr) = 1 'ON
    End Select
End Sub

Sub FlashAR(nr, ramp, a, b, c)                                   'used for reflections when there is no off ramp
    Select Case FadingLevel(nr)
        Case 2:ramp.alpha = 0:FadingLevel(nr) = 0                'Off
        Case 3:ramp.image = c:FadingLevel(nr) = 2                'fading...
        Case 4:ramp.image = b:FadingLevel(nr) = 3                'fading...
        Case 5:ramp.image = a:ramp.alpha = 1:FadingLevel(nr) = 1 'ON
    End Select
End Sub

Sub FlashARm(nr, ramp, a, b, c)
    Select Case FadingLevel(nr)
        Case 2:ramp.alpha = 0
        Case 3:ramp.image = c
        Case 4:ramp.image = b
        Case 5:ramp.image = a:ramp.alpha = 1
    End Select
End Sub

Sub NFadeAR(nr, ramp, a, b)
    Select Case FadingLevel(nr)
        Case 4:ramp.image = b:FadingLevel(nr) = 0 'off
        Case 5:ramp.image = a:FadingLevel(nr) = 1 'on
    End Select
End Sub

Sub NFadeARm(nr, ramp, a, b)
    Select Case FadingLevel(nr)
        Case 4:ramp.image = b
        Case 5:ramp.image = a
    End Select
End Sub

Sub MNFadeAR(nr, ramp, a, b, c)
    Select Case FadingLevel(nr)
        Case 4:ramp.image = b:FadingLevel(nr) = 0:SetLamp c, FadingLevel(c) 'off
        Case 5:ramp.image = a:FadingLevel(nr) = 1                           'on
    End Select
End Sub

Sub MNFadeARm(nr, ramp, a, b, c)
    Select Case FadingLevel(nr)
        Case 4:ramp.image = b:SetLamp c, FadingLevel(c) 'off
        Case 5:ramp.image = a                           'on
    End Select
End Sub

' Flashers using PRIMITIVES
' pri is the name of the primitive
' a,b,c,d are the images used for on...off

Sub FadePri(nr, pri, a, b, c, d)
    Select Case FadingLevel(nr)
        Case 2:pri.image = d:FadingLevel(nr) = 0 'Off
        Case 3:pri.image = c:FadingLevel(nr) = 2 'fading...
        Case 4:pri.image = b:FadingLevel(nr) = 3 'fading...
        Case 5:pri.image = a:FadingLevel(nr) = 1 'ON
    End Select
End Sub

'Fades primitives, 3 images

Sub FadePri3m(nr, pri, group)
    Select Case FadingLevel(nr)
        Case 3:pri.image = group(2)
        Case 4:pri.image = group(1)
        Case 5:pri.image = group(0)
    End Select
End Sub

Sub FadePri3(nr, pri, group)
    Select Case FadingLevel(nr)
        Case 3:pri.image = group(2):FadingLevel(nr) = 0 'Off
        Case 4:pri.image = group(1):FadingLevel(nr) = 3 'fading...
        Case 5:pri.image = group(0):FadingLevel(nr) = 1 'ON
    End Select
End Sub


'Fades primitives, 4 images

Sub FadePri4m(nr, pri, group)
    Select Case FadingLevel(nr)
		Case 2:pri.image = group(3)
        Case 3:pri.image = group(2)
        Case 4:pri.image = group(1)
        Case 5:pri.image = group(0)
    End Select
End Sub


Sub FadePri4(nr, pri, group)
    Select Case FadingLevel(nr)
		Case 2:pri.image = group(3):FadingLevel(nr) = 0 'Off
        Case 3:pri.image = group(2):FadingLevel(nr) = 2 'fading...
        Case 4:pri.image = group(1):FadingLevel(nr) = 3 'fading...
        Case 5:pri.image = group(0):FadingLevel(nr) = 1 'ON
    End Select
End Sub


Sub FadePriC(nr, pri, a, b, c, d)
    Select Case FadingLevel(nr)
        Case 2:For each xx in pri:xx.image = d:Next:FadingLevel(nr) = 0 'Off
        Case 3:For each xx in pri:xx.image = c:Next:FadingLevel(nr) = 2 'fading...
        Case 4:For each xx in pri:xx.image = b:Next:FadingLevel(nr) = 3 'fading...
        Case 5:For each xx in pri:xx.image = a:Next:FadingLevel(nr) = 1 'ON
    End Select
End Sub

Sub FadePrih(nr, pri, a, b, c, d)
    Select Case FadingLevel(nr)
        Case 2:pri.image = d:SetFlash nr, 0:FadingLevel(nr) = 0 'Off
        Case 3:pri.image = c:FadingLevel(nr) = 2 'fading...
        Case 4:pri.image = b:FadingLevel(nr) = 3 'fading...
        Case 5:pri.image = a:SetFlash nr, 1:FadingLevel(nr) = 1 'ON
    End Select
End Sub

Sub FadePrim(nr, pri, a, b, c, d)
    Select Case FadingLevel(nr)
        Case 2:pri.image = d
        Case 3:pri.image = c
        Case 4:pri.image = b
        Case 5:pri.image = a
    End Select
End Sub

Sub NFadePri(nr, pri, a, b)
    Select Case FadingLevel(nr)
        Case 4:pri.image = b:FadingLevel(nr) = 0 'off
        Case 5:pri.image = a:FadingLevel(nr) = 1 'on
    End Select
End Sub

Sub NFadePrim(nr, pri, a, b)
    Select Case FadingLevel(nr)
        Case 4:pri.image = b
        Case 5:pri.image = a
    End Select
End Sub

'Fade a collection of lights

Sub FadeLCo(nr, a, b) 'fading collection of lights
    Dim obj
    Select Case FadingLevel(nr)
        Case 2:vpmSolToggleObj b, Nothing, 0, 0:FadingLevel(nr) = 0
        Case 3:vpmSolToggleObj b, Nothing, 0, 1:FadingLevel(nr) = 2
        Case 4:vpmSolToggleObj a, Nothing, 0, 0:FadingLevel(nr) = 3
        Case 5:vpmSolToggleObj a, Nothing, 0, 1:FadingLevel(nr) = 1
    End Select
End Sub


' *********************************************************************
'                      Supporting Ball & Sound Functions
' *********************************************************************

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
    Vol = Csng(BallVel(ball) ^2 / 2000)
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

Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
    Pitch = BallVel(ball) * 20
End Function

Function BallVel(ball) 'Calculates the ball speed
    BallVel = INT(SQR((ball.VelX ^2) + (ball.VelY ^2) ) )
End Function

'*****************************************
'      JP's VP10 Rolling Sounds
'*****************************************

Const tnob = 15 ' total number of balls
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
        If BallVel(BOT(b) ) > 1 AND BOT(b).z < 30 Then
            rolling(b) = True
            PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) ), Pan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0
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
    If ball1.id = 666 OR ball2.id = 666 Then
    PlaySound("rubber_hit_3"), 0, Csng(velocity) ^2 / 2000, Pan(ball1), 0, Pitch(ball1), 0, 0
    else
	PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 2000, Pan(ball1), 0, Pitch(ball1), 0, 0
    End If
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
	PlaySound "pinhit_low", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0
End Sub

Sub Targets_Hit (idx)
	PlaySound "target", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0
End Sub

Sub Metals_Thin_Hit (idx)
	PlaySound "metalhit_thin", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
End Sub

Sub Metals_Medium_Hit (idx)
	PlaySound "metalhit_medium", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
End Sub

Sub Metals2_Hit (idx)
	PlaySound "metalhit2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
End Sub

Sub Gates_Hit (idx)
	PlaySound "gate4", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
End Sub

Sub Spinner_Spin
	PlaySound "fx_spinner",0,.25,0,0.25
End Sub

Sub Rubbers_Hit(idx)
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 20 then 
		PlaySound "fx_rubber2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
	End if
	If finalspeed >= 6 AND finalspeed <= 20 then
 		RandomSoundRubber()
 	End If
End Sub

Sub Posts_Hit(idx)
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 16 then 
		PlaySound "fx_rubber2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
	End if
	If finalspeed >= 6 AND finalspeed <= 16 then
 		RandomSoundRubber()
 	End If
End Sub

Sub RandomSoundRubber()
	Select Case Int(Rnd*3)+1
		Case 1 : PlaySound "rubber_hit_1", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
		Case 2 : PlaySound "rubber_hit_2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
		Case 3 : PlaySound "rubber_hit_3", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
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
		Case 1 : PlaySound "flip_hit_1", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
		Case 2 : PlaySound "flip_hit_2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
		Case 3 : PlaySound "flip_hit_3", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
	End Select
End Sub

' Get angle
Dim Xin, Yin, rAngle, Radit, wAngle, Pi
Pi = Round(4 * Atn(1), 6) '3.1415926535897932384626433832795

Sub GetAngle(Xin, Yin, wAngle)
    If Sgn(Xin) = 0 Then
        If Sgn(Yin) = 1 Then rAngle = 3 * Pi / 2 Else rAngle = Pi / 2
        If Sgn(Yin) = 0 Then rAngle = 0
        Else
            rAngle = atn(- Yin / Xin)
    End If
    If sgn(Xin) = -1 Then Radit = Pi Else Radit = 0
    If sgn(Xin) = 1 and sgn(Yin) = 1 Then Radit = 2 * Pi
    wAngle = round((Radit + rAngle), 4)
End Sub

'Rotation of the small saucers
Dim uforotny
uforotny = 0
Sub SmallUfoTimer_Timer()
uforotny = (uforotny +8) MOD 360
ufo2.RotY = -uforotny
ufo2d.RotZ = -uforotny	
End Sub

Sub Mars_Hit()
If Mars_Mod = 1 Then can2.visible = 1:UFO_Mars.enabled = 1:End If
End Sub

Dim mspin
Sub UFO_Mars_Timer()
If mspin = 3000 Then
mspin = 0
me.enabled = 0
can2.visible = 0
End If
can2.Rotz = can2.Rotz + 1
mspin = mspin + 1
End Sub

Sub Table1_exit()
	Controller.Pause = False
	Controller.Stop
End Sub

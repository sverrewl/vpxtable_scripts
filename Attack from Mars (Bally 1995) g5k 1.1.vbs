' _______ __________________ _______  _______  _       
'(  ___  )\__   __/\__   __/(  ___  )(  ____ \| \    /\
'| (   ) |   ) (      ) (   | (   ) || (    \/|  \  / /
'| (___) |   | |      | |   | (___) || |      |  (_/ / 
'|  ___  |   | |      | |   |  ___  || |      |   _ (  
'| (   ) |   | |      | |   | (   ) || |      |  ( \ \ 
'| )   ( |   | |      | |   | )   ( || (____/\|  /  \ \
'|/     \|   )_(      )_(   |/     \|(_______/|_/    \/
'                                                      
' _______  _______  _______  _______    _______  _______  _______  _______ 
'(  ____ \(  ____ )(  ___  )(       )  (       )(  ___  )(  ____ )(  ____ \
'| (    \/| (    )|| (   ) || () () |  | () () || (   ) || (    )|| (    \/
'| (__    | (____)|| |   | || || || |  | || || || (___) || (____)|| (_____ 
'|  __)   |     __)| |   | || |(_)| |  | |(_)| ||  ___  ||     __)(_____  )
'| (      | (\ (   | |   | || |   | |  | |   | || (   ) || (\ (         ) |
'| )      | ) \ \__| (___) || )   ( |  | )   ( || )   ( || ) \ \__/\____) |
'|/       |/   \__/(_______)|/     \|  |/     \||/     \||/   \__/\_______)
'                                                                          
'
'
'
'  ______________               ___ __  _           
'  / ____/ ____/ /__   ___  ____/ (_) /_(_)___  ____ 
' / / __/___ \/ //_/  / _ \/ __  / / __/ / __ \/ __ \
'/ /_/ /___/ / ,<    /  __/ /_/ / / /_/ / /_/ / / / /
'\____/_____/_/|_|   \___/\__,_/_/\__/_/\____/_/ /_/ 
'


'//////////////////////////////////////////////////////////////////////
' Update VPINMAME.DLL to support WPC-95 FASTFLIPS (SAMBuild r4550 or newer)
'http://vpuniverse.com/forums/topic/3461-sambuild31-beta-thread/
'//////////////////////////////////////////////////////////////////////


' Attack from Mars - IPDB No. 3781
' © Bally Williams 1995
' Design by Brian Eddy
' VPX recreation by g5k
' Thanks to:
' Tom Tower: Ramp, Main Saucer and Alien primitives
' DJRobX: Fastflips WPC95, SSF and sound sweeten, develop assistance
' JimmyFingers for passing on lots of helpful knowledge to do with GI
' Previous AFM VPX authors who's work helped guide the build on this one and to the components that are carried over
' To those in the Monster Bash Pincab and VP Junkies community that helped contributed artwork, photos and the banter
' The VPDevTeam and authors who have contributed to the projects and components that get passed on from table to table 
' I'm sure there are primitives, scripts, sounds and so on that have been by others not mentioned above, thank you for your contributions!
' Thanks to Monsterbash Pincab for distribution and providing such a great resource for the community


' AFM Gameplay rules
' http://pinball.org/rules/attackfrommars.html


'//////////////////////////////////////////////////////////////////////
'// LAYER 01 = STANDARD VP OBJECTS 
'// LAYER 02 = MAIN SAUCER + DIVERTER 
'// LAYER 03 = INSERT LIGHTS 
'// LAYER 04 = RED FLASHERS 
'// LAYER 05 = BLUE/GREEN FLASHERS 
'// LAYER 06 = G5k 3D 
'// LAYER 07 = GI - Lower Zone  
'// LAYER 08 = GI - Mid Zone  
'// LAYER 09 = GI - Upper Zone  
'// LAYER 10 = INSERTs Spill to other objects (Cheap SSR) 
'// LAYER 11 = INSERTs LIGHT GLOW/BALL REFLECT
'//////////////////////////////////////////////////////////////////////


'NOTE for 120hz users, select the object on the table, "LAMPTIMER" on layer 1 and change the value to 16 which will force the speed to match how it looks at 60hz
'-1 (this is default and correct for 60hz)



Option Explicit
Randomize


'///////////////////////////////////////////////////////////////////////////////////////////////////
' _       ______  ______     ____  ______   _________   ___________   ________    ________  _____    
'| |     / / __ \/ ____/    / __ \/ ____/  / ____/   | / ___/_  __/  / ____/ /   /  _/ __ \/ ___/    
'| | /| / / /_/ / /  ______/ /_/ /___ \   / /_  / /| | \__ \ / /    / /_  / /    / // /_/ /\__ \     
'| |/ |/ / ____/ /__/_____/\__, /___/ /  / __/ / ___ |___/ // /    / __/ / /____/ // ____/___/ /     
'|__/|__/_/    \____/     /____/_____/  /_/   /_/  |_/____//_/    /_/   /_____/___/_/    /____/      
'                                                                                                    
'
'///////////////////////////////////////////////////////////////////////////////////////////////////
'// FAST FLIPS (DJRobX) - needs vpinmame.dll
'// Available in SAMupdate 4550 or newer 
'// http://vpuniverse.com/forums/topic/3461-sambuild31-beta-thread/
'// If you don't want to update change UseSolendoids below = 1
'//////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////
'// UseSolenoids = [1 or 2] 1 = Standard flipper latency, 2 = FastFlips 

Const UseSolenoids = 2

'//////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////


'//////////////////////////////////////////////////////////////////////
'Options to disable if your system is not coping
'1=YES/ON     0=NO/OFF

'If you are experiencing performance issues, the first thing you should do is turn off playfield reflections in table options/playfield graphics.
'Next I recommend changing STEELWALLS primitive to static rendering (large texture, but mostly obscured by other things on table)

'PRELOAD ALL THE FLASHERS ON STARTUP
'NOTE: If saucer hesitates on first hit after lowering the bank
'you can use this and might help performance overall slightly.   Causes the table to do a flash sequence at startup.
Const PreloadMe = 1

'OPTIONAL COOL STUFF
'
'If you don't like things that are cool disable flasher effects below
'Might improve performance too :)
' 1 = ON (Default awesome), 0 = OFF (Lame)
' Additional performance tweaks:   
' 
Const AlienFlashers = 1 'New to v1.1
Const SideWallFlashers = 1 'These add a lot visually but are also taxing
Const ChromeRailFlashers = 1 'New to v1.1.  May be better to change CHROMERAILS primitive to static rendering, will give more performance and improve AA quality (Note: If static though you do not get flashing effects on the rails)
Const UFOFlashers = 1
Const PlasticsTopFlasher = 1 'New to v1.1
Const ApronGIProper = 1 'This would be the first thing to lose if you are trying to add performance as difference is minor.  May also change APRON primitive to static rendering.







'//////////////////////////////////////////////////////////////////////

Const UseLamps = 0
Const UseSync = 1
Const HandleMech = 0
Const SCoin = "fx_Coin"
Const UseVPMModSol = 1


'//////////////////////////////////////////////////////////////////////
'// DOF CONTROLLER INIT 
'//////////////////////////////////////////////////////////////////////

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

'//////////////////////////////////////////////////////////////////////
'// VARIABLES 
'//////////////////////////////////////////////////////////////////////

Dim bsTrough, plungerIM, Ballspeed, bsL, bsR, aBall, bBall, dtDrop, currGITop, currGIMid, currGIBot
Dim DesktopMode: DesktopMode = AFM.ShowDT
Dim UseVPMDMD:UseVPMDMD = DesktopMode			'shows VPX internal DMD


'//////////////////////////////////////////////////////////////////////
'// VPM INIT 
'//////////////////////////////////////////////////////////////////////

Const BallSize = 50
Const Ballmass = 1.7

LoadVPM "01560000", "WPC.VBS", 3.46

' Wob: Needed for Fast Flips
NoUpperRightFlipper
NoUpperLeftFlipper

Sub LoadVPM(VPMver, VBSfile, VBSver)
On Error Resume Next
If ScriptEngineMajorVersion < 5 Then MsgBox "VB Script Engine 5.0 or higher required"
ExecuteGlobal GetTextFile(VBSfile)
If Err Then MsgBox "Unable to open " & VBSfile & ". Ensure that it is in the same folder as this table. " & vbNewLine & Err.Description
Set Controller = CreateObject("VPinMAME.Controller")
If Err Then MsgBox "Can't Load VPinMAME." & vbNewLine & Err.Description
If VPMver > "" Then If Controller.Version < VPMver Or Err Then MsgBox "VPinMAME ver " & VPMver & " required."
If VPinMAMEDriverVer < VBSver Or Err Then MsgBox VBSFile & " ver " & VBSver & " or higher required."
On Error Goto 0

End Sub



Set GiCallback2 = GetRef("UpdateGI")

'//////////////////////////////////////////////////////////////////////
'// GAME NAME 
'//////////////////////////////////////////////////////////////////////

Const cGameName = "afm_113b"


'//////////////////////////////////////////////////////////////////////
'// STARTUP SETTINGS 
'//////////////////////////////////////////////////////////////////////

SW45.isdropped=1:SW46.isdropped=1:SW47.isdropped=1

RightSlingRubber_A.visible=0
LeftSlingRubber_A.visible=0
SLINGL.visible=0
SLINGR.visible=0





If ApronGIProper = 1 then APRON_GIoff.visible = 1 else APRON_GIoff.visible = 0

'Show Desktop components and adjust sidewalls
	If DesktopMode Then
	Ufo.Z=20
	LED1.Z=20
LED10.Z=20
LED11.Z=20
LED12.Z=20
LED13.Z=20
LED14.Z=20
LED15.Z=20
LED16.Z=20
LED2.Z=20
LED3.Z=20
LED4.Z=20
LED5.Z=20
LED6.Z=20
LED7.Z=20
LED8.Z=20
LED9.Z=20
METALBRUSHED.Z=20
'		SIDERAILS_B.visible=1
'		CABINETSIDES.SIZE_Y=21
'		CABINETSIDES.X=0
'		CABINETSIDES.Z=0
'		CABINETSIDES_GIOFF.SIZE_Y=21
'		CABINETSIDES_GIOFF.X=0
'		CABINETSIDES_GIOFF.Z=0

	Else
		SIDERAILS_B.visible=0
	End if

CENTREBANK.Z=-76:SW45P.Z=-76:SW46P.Z=-76:SW47P.Z=-76

Diverter8.isdropped=1:Diverter7.isdropped=1:Diverter6.isdropped=1:Diverter5.isdropped=1

Diverter4.isdropped=1:Diverter3.isdropped=1:Diverter2.isdropped=1:Diverter1.isdropped=1

DivPos=0



'//////////////////////////////////////////////////////////////////////
'// TABLE INIT 
'//////////////////////////////////////////////////////////////////////


Sub AFM_Init
vpmInit Me
With Controller
.GameName = cGameName
If Err Then MsgBox "Can't start Game " & cGameName & vbNewLine & Err.Description:Exit Sub
.SplashInfoLine = "Attack from Mars" & vbNewLine & "g5k Edition"
.Games(cGameName).Settings.Value("rol") = 0
.HandleKeyboard = 0
.ShowTitle = 0
.ShowDMDOnly = 1
.ShowFrame = 0
.HandleMechanics = 0
.Hidden = 0
On Error Resume Next
.Run GetPlayerHWnd
If Err Then MsgBox Err.Description
On Error Goto 0
.Switch(22) = 1 'close coin door
.Switch(24) = 1 'and keep it close
End With

'//////////////////////////////////////////////////////////////////////
'// PINMAME TIMER 
'//////////////////////////////////////////////////////////////////////

PinMAMETimer.Interval = PinMAMEInterval
PinMAMETimer.Enabled = 1

'//////////////////////////////////////////////////////////////////////
'// BALL TROUGH 
'//////////////////////////////////////////////////////////////////////

Set bsTrough = New cvpmTrough
    With bsTrough
        .size = 4
        .initSwitches Array(32, 33, 34, 35)
        .Initexit BallRelease, 90, 4
        .InitExitSounds SoundFX("BallRelease",DOFContactors), SoundFX("BallRelease",DOFContactors)
        .Balls = 4
    End With


'//////////////////////////////////////////////////////////////////////
'// IMPULSE PLUNGER 
'//////////////////////////////////////////////////////////////////////

Const IMPowerSetting = 60 'Plunger Power
Const IMTime = 0.6        ' Time in seconds for Full Plunge
Set plungerIM = New cvpmImpulseP
With plungerIM
.InitImpulseP swplunger, IMPowerSetting, IMTime
.Random 3
.switch 18
.InitExitSnd SoundFX("ShooterLane",DOFContactors), ""
.CreateEvents "plungerIM"
End With
TBPos=28:TBTimer.Enabled=0:TBDown=1:Controller.Switch(66) = 1:Controller.Switch(67) = 0


'//////////////////////////////////////////////////////////////////////
'// VUK 
'//////////////////////////////////////////////////////////////////////
 
  ' Left hole
    Set bsL = New cvpmTrough
    With bsL
        .size = 1
        .initSwitches Array(36)
        .Initexit sw36, 180, 0
        .InitExitSounds "", SoundFX("VUKOut",DOFContactors)
        .InitExitVariance 3, 2
    End With
  
' Set bsL = New cvpmBallStack
' With bsL
' .InitSw 0, 36, 0, 0, 0, 0, 0, 0
' .InitKick sw36, 180, 0
' .InitExitSnd SoundFX("VUKOut",DOFContactors), ""
' .KickForceVar = 3
' End With

'//////////////////////////////////////////////////////////////////////
'// SCOOP 
'//////////////////////////////////////////////////////////////////////

 Set bsR = New cvpmTrough
    With bsR
        .size = 4
        .initSwitches Array(37)
        .Initexit sw37, 198, 30
        .InitExitSounds "", SoundFX("ScoopExit",DOFContactors)
        .InitExitVariance 2, 2
        .MaxBallsPerKick = 1
		'.KickZ = 0.4
		'.KickAngleVar = 1
    End With

' Set bsR = New cvpmBallStack
' With bsR
' .InitSw 0, 37, 0, 0, 0, 0, 0, 0
' .InitKick SW37, 198, 30
' .KickZ = 0.4
' .InitExitSnd SoundFx("ScoopExit",DOFContactors), ""
' .KickAngleVar = 1
' .KickBalls = 1
' End With

'//////////////////////////////////////////////////////////////////////
'// DROPTARGET 
'//////////////////////////////////////////////////////////////////////
 Set dtDrop = New cvpmDropTarget
 With dtDrop
 .InitDrop sw77, 77
 .initsnd "", ""
 End With

'//////////////////////////////////////////////////////////////////////
'// NUDGING 
'//////////////////////////////////////////////////////////////////////

vpmNudge.TiltSwitch = 14
vpmNudge.Sensitivity = 5.35
vpmNudge.TiltObj = Array(LeftBumper, BottomBumper, RightBumper, LeftSlingshot, RightSlingshot)

'//////////////////////////////////////////////////////////////////////
'// UPDATE GI 
'//////////////////////////////////////////////////////////////////////

UpdateGI 0, 1:UpdateGI 1, 1:UpdateGI 2, 1



'//////////////////////////////////////////////////////////////////////
'// END OF TABLE INIT 
'//////////////////////////////////////////////////////////////////////
End Sub




'//////////////////////////////////////////////////////////////////////
'// CONTROLLER PAUSE 
'//////////////////////////////////////////////////////////////////////

Sub AFM_Paused:Controller.Pause = 1:End Sub
Sub AFM_unPaused:Controller.Pause = 0:End Sub

'//////////////////////////////////////////////////////////////////////
'// DRAIN HIT 
'//////////////////////////////////////////////////////////////////////

Sub Drain_Hit:PlaySoundAtVol "Drain",Drain,.8:bsTrough.AddBall Me:End Sub
Sub Drain1_Hit:PlaySoundAtVol "Drain",Drain,.8:bsTrough.AddBall Me:End Sub
Sub Drain2_Hit:PlaySoundAtVol "Drain",Drain,.8:bsTrough.AddBall Me:End Sub

'//////////////////////////////////////////////////////////////////////
'// VUK HIT 
'//////////////////////////////////////////////////////////////////////

Sub sw36a_Hit:PlaySoundAtVol "VUKEnter",sw36a,1:bsL.AddBall Me:End Sub

'//////////////////////////////////////////////////////////////////////
'// SCOOP HIT 
'//////////////////////////////////////////////////////////////////////

Sub sw37a_Hit:PlaySoundAtVol "ScoopBack",sw37a,1:bsR.AddBall Me:End Sub

Sub sw37_Hit
PlaySoundAtVol "ScoopFront",sw37,1
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

'//////////////////////////////////////////////////////////////////////
'// SAUCER DROPHOLE HIT 
'//////////////////////////////////////////////////////////////////////

Sub sw78_Hit:PlaySoundAtVol "Drophole",sw78,1
Set aBall = ActiveBall
Me.TimerEnabled =1
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

'//////////////////////////////////////////////////////////////////////
'// SAUCER TARGET HIT 
'//////////////////////////////////////////////////////////////////////

Sub sw77_Hit:dtDrop.Hit 1:End Sub

'//////////////////////////////////////////////////////////////////////
'// KEYS 
'//////////////////////////////////////////////////////////////////////

Sub AFM_KeyDown(ByVal Keycode)
If keycode = PlungerKey Then Controller.Switch(11) = 1:End If
If keycode = CenterTiltKey Then Nudge 0, 5
If vpmKeyDown(keycode) Then Exit Sub
End Sub

Sub AFM_KeyUp(ByVal Keycode)
If keycode = PlungerKey Then Controller.Switch(11) = 0
If vpmKeyUp(keycode) Then Exit Sub
End Sub

'//////////////////////////////////////////////////////////////////////
'// SOLENOID CALLS 
'//////////////////////////////////////////////////////////////////////

SolCallback(1) = "Auto_Plunger"
SolCallback(2) = "SolRelease"
SolCallback(3) = "bsL.SolOut"
SolCallback(4) = "bsR.SolOut"
SolCallBack(5) = "Alien1Move"
SolCallBack(6) = "Alien2Move"
SolCallback(7) = "vpmSolSound SoundFX(""fx_Knocker"",DOFKnocker),"
SolCallBack(8) = "Alien3Move"
SolCallBack(14) = "Alien4Move"
SolCallBack(15) = "SolUfoShake"
SolCallback(16) = "ResetUFODrop"
SolModCallback(17) = "setmodlamp 117,"
SolModCallback(18) = "setmodlamp 118,"
SolModCallback(19) = "setmodlamp 119,"
SolModCallBack(20) = "setmodlamp 120,"
SolModCallback(21) = "setmodlamp 121,"
SolModCallback(22) = "setmodlamp 122,"
SolModCallback(23) = "setmodlamp 123,"
SolCallBack(24) = "TBMove"
SolModCallBack(25) = "setmodlamp 125,"
SolModCallBack(26) = "setmodlamp 126,"
SolModCallBack(27) = "setmodlamp 127,"
SolModCallBack(28) = "setmodlamp 128,"
SolCallback(33) = "vpmSolGate RGate,false,"
SolCallback(34) = "vpmSolGate LGate,false,"
SolCallback(36) = "DivMove"
SolModCallback(39) = "setmodlamp 123,"
SolModCallback(43) = "setmodlamp 130,"
SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

'//////////////////////////////////////////////////////////////////////
'// FLIPPERS 
'//////////////////////////////////////////////////////////////////////

Sub SolLFlipper(Enabled)
If UseSolenoids = 2 then Controller.Switch(swULFlip)=Enabled
If Enabled Then
PlaySoundAtVol SoundFX("FlipperUpRight",DOFContactors),LeftFlipper,3
LeftFlipper.RotateToEnd
Else
PlaySoundAtVol SoundFX("FlipperDownLeft",DOFContactors),LeftFlipper,.2
LeftFlipper.RotateToStart
End If
End Sub

Sub SolRFlipper(Enabled)
if UseSolenoids = 2 then Controller.Switch(swURFlip)=Enabled
If Enabled Then
PlaySoundAtVol SoundFX("FlipperUpRight",DOFContactors),RightFlipper,2
RightFlipper.RotateToEnd
Else
PlaySoundAtVol SoundFX("FlipperDownRight",DOFContactors),RightFlipper,.2
RightFlipper.RotateToStart
End If
End Sub

'//////////////////////////////////////////////////////////////////////
'// SOLENOID SUBS 
'//////////////////////////////////////////////////////////////////////

Sub Auto_Plunger(Enabled)
If Enabled Then
PlungerIM.AutoFire
End If
End Sub

Sub SolRelease(Enabled)
If Enabled And bsTrough.Balls> 0 Then
vpmTimer.PulseSw 31
bsTrough.ExitSol_On
End If
End Sub

'//////////////////////////////////////////////////////////////////////
'// POP BUMPERS 
'//////////////////////////////////////////////////////////////////////

Sub LeftBumper_Hit:vpmTimer.PulseSw 53:PlaySoundAtBumperVol "LeftBumper", LeftBumper,3
End Sub

Sub BottomBumper_Hit:vpmTimer.PulseSw 54:PlaySoundAtBumperVol "BottomBumper", BottomBumper,3
End Sub

Sub RightBumper_Hit:vpmTimer.PulseSw 55:PlaySoundAtBumperVol "RightBumper", RightBumper,3
End Sub

'//////////////////////////////////////////////////////////////////////
'// SLINGSHOTS 
'//////////////////////////////////////////////////////////////////////

Sub LeftSlingShot_Slingshot:LeftSlingRubber_A.visible = 1:SLINGL.visible = 1:PlaySound SoundFX("LeftSlingshot",DOFContactors):vpmTimer.PulseSw 51:LeftSlingshot.timerenabled = 1:End Sub
Sub LeftSlingshot_timer:LeftSlingRubber_A.visible = 0:SLINGL.visible = 0:LeftSlingshot.timerenabled= 0:End Sub

Sub RightSlingShot_Slingshot:RightSlingRubber_A.visible = 1:SLINGR.visible = 1:PlaySound SoundFX("RightSlingshot",DOFContactors):vpmTimer.PulseSw 52:Rightslingshot.timerenabled = 1:End Sub
Sub Rightslingshot_timer:RightSlingRubber_A.visible = 0:SLINGR.visible = 0:Rightslingshot.timerenabled= 0:End Sub



'//////////////////////////////////////////////////////////////////////
'// MARTIAN TARGETS 
'//////////////////////////////////////////////////////////////////////


'// "M"ARTIAN TARGET //////////////////////////////////////////////////

Sub SW56_Hit
	If SW56P.ObjRotY = 0 Then 
		vpmTimer.PulseSw 56:SW56P.ObjRotY=1:SW56P1.ObjRotY=1:Me.TimerEnabled=1
		PlaySound SoundFX("TargetHit",DOFContactors),0,Vol(ActiveBall),Pan(ActiveBall),0,0,1,0,AudioFade(ActiveBall)
	End If
End Sub

Sub SW56_Timer:SW56P.ObjRotY=0:SW56P1.ObjRotY=0:Me.TimerEnabled=0:End Sub


'// M"A"RTIAN TARGET //////////////////////////////////////////////////

Sub SW57_Hit
	If SW57P.ObjRotY = 0 Then 
		vpmTimer.PulseSw 57:SW57P.ObjRotY=1:SW57P1.ObjRotY=1:Me.TimerEnabled = 1
		PlaySound SoundFX("TargetHit",DOFContactors),0,Vol(ActiveBall),Pan(ActiveBall),0,0,1,0,AudioFade(ActiveBall)
	End If
End Sub

Sub SW57_Timer:SW57P.ObjRotY=0:SW57P1.ObjRotY=0:Me.TimerEnabled=0:End Sub

'// MA"R"TIAN TARGET //////////////////////////////////////////////////

Sub SW58_Hit
	If SW58P.ObjRotY = 0 Then
		vpmTimer.PulseSw 58:SW58P.ObjRotY=-1:SW58P1.ObjRotY=-1:Me.TimerEnabled = 1
		PlaySound SoundFX("TargetHit",DOFContactors),0,Vol(ActiveBall),Pan(ActiveBall),0,0,1,0,AudioFade(ActiveBall)
	End If
End Sub

Sub SW58_Timer:SW58P.ObjRotY=0:SW58P1.ObjRotY=0:Me.TimerEnabled=0:End Sub


'// MAR"T"IAN TARGET //////////////////////////////////////////////////


Sub SW43_Hit
	If SW43P.ObjRotX = 0 Then
		vpmTimer.PulseSw 43:SW43P.ObjRotX=0.35:SW43P1.ObjRotX=0.35:Me.TimerEnabled=1
		PlaySound SoundFX("TargetHit",DOFContactors),0,Vol(ActiveBall),Pan(ActiveBall),0,0,1,0,AudioFade(ActiveBall)
	End If
End Sub

Sub SW43_Timer:SW43P.ObjRotX=0:SW43P1.ObjRotX=0:Me.TimerEnabled=0:End Sub


'// MART"I"AN TARGET //////////////////////////////////////////////////

Sub SW44_Hit
	If SW44P.ObjRotY = 0 Then
		vpmTimer.PulseSw 44:SW44P.ObjRotX = 0.35:SW44P1.ObjRotX = 0.35:Me.TimerEnabled = 1
		PlaySound SoundFX("TargetHit",DOFContactors),0,Vol(ActiveBall),Pan(ActiveBall),0,0,1,0,AudioFade(ActiveBall)
	End If
End Sub

Sub SW44_Timer:SW44P.ObjRotX = 0:SW44P1.ObjRotX = 0:Me.TimerEnabled = 0:End Sub



'// MARTI"A"N TARGET //////////////////////////////////////////////////

Sub SW41_Hit
	If SW41P.ObjRotY = 0 Then 
		vpmTimer.PulseSw 41:SW41P.ObjRotY = 1:SW41P1.ObjRotY = 1:Me.TimerEnabled = 1
		PlaySound SoundFX("TargetHit",DOFContactors),0,Vol(ActiveBall),Pan(ActiveBall),0,0,1,0,AudioFade(ActiveBall)
	End If
End Sub

Sub SW41_Timer:SW41P.ObjRotY = 0:SW41P1.ObjRotY = 0:Me.TimerEnabled = 0:End Sub



'// MARTIA"N" TARGET //////////////////////////////////////////////////

Sub SW42_Hit
	If SW42P.ObjRotY = 0 Then
		vpmTimer.PulseSw 42:SW42P.ObjRotY = 1:SW42P1.ObjRotY = 1:Me.TimerEnabled = 1
		PlaySound SoundFX("TargetHit",DOFContactors),0,Vol(ActiveBall),Pan(ActiveBall),0,0,1,0,AudioFade(ActiveBall)
	End If
End Sub

Sub SW42_Timer:SW42P.ObjRotY = 0:SW42P1.ObjRotY = 0:Me.TimerEnabled = 0:End Sub

'//////////////////////////////////////////////////////////////////////
'// SAUCER TARGETS 
'//////////////////////////////////////////////////////////////////////

Sub SW75_Hit
	If SW75P.ObjRotY = 0 Then
		vpmTimer.PulseSw 75:SW75P.ObjRotY = 1:Me.TimerEnabled = 1
		PlaySound SoundFX("TargetHit",DOFContactors),0,Vol(ActiveBall),Pan(ActiveBall),0,0,1,0,AudioFade(ActiveBall)
	End If
End Sub

Sub SW75_Timer:SW75P.ObjRotY = 0:Me.TimerEnabled = 0:End Sub

Sub SW76_Hit
	If SW76P.ObjRotY = 0 Then
		vpmTimer.PulseSw 76:SW76P.ObjRotY = 1:Me.TimerEnabled = 1
		PlaySound SoundFX("TargetHit",DOFContactors),0,Vol(ActiveBall),Pan(ActiveBall),0,0,1,0,AudioFade(ActiveBall)
	End If
End Sub

Sub SW76_Timer:SW76P.ObjRotY = 0:Me.TimerEnabled = 0:End Sub

'//////////////////////////////////////////////////////////////////////
'// TARGETBANK TARGETS 
'//////////////////////////////////////////////////////////////////////

Sub SW45_Hit
	If SW45P.ObjRotX = 0 Then
		vpmTimer.PulseSw 45:SW45P.ObjRotX = 0.35:Me.TimerEnabled = 1
		PlaySound SoundFX("TargetHit",DOFContactors),0,Vol(ActiveBall),Pan(ActiveBall),0,0,1,0,AudioFade(ActiveBall)
	End If
	If Ballresting = True Then
		DPBall.VelY = ActiveBall.VelY * 3
	End If
End Sub

Sub SW45_Timer:SW45P.ObjRotX = 0:Me.TimerEnabled = 0:End Sub

Sub SW46_Hit
	If SW46P.ObjRotX = 0 Then
		vpmTimer.PulseSw 46:SW46P.ObjRotX = 0.35:Me.TimerEnabled = 1
		PlaySound SoundFX("TargetHit",DOFContactors),0,Vol(ActiveBall),Pan(ActiveBall),0,0,1,0,AudioFade(ActiveBall)
	End If
	If Ballresting = True Then
		DPBall.VelY = ActiveBall.VelY * 3
	End If
End Sub

Sub SW46_Timer:SW46P.ObjRotX = 0:Me.TimerEnabled = 0:End Sub

Sub SW47_Hit
	If SW47P.ObjRotX = 0 Then
		vpmTimer.PulseSw 47:SW47P.ObjRotX = 0.35:Me.TimerEnabled = 1
		PlaySound SoundFX("TargetHit",DOFContactors),0,Vol(ActiveBall),Pan(ActiveBall),0,0,1,0,AudioFade(ActiveBall)
	End If
	If Ballresting = True Then
		DPBall.VelY = ActiveBall.VelY * 3
	End If
End Sub

Sub SW47_Timer:SW47P.ObjRotX = 0:Me.TimerEnabled = 0:End Sub


Sub sw77_Dropped:dtDrop.Hit 1:End Sub

Sub ResetUFODrop(enabled)
	dtDrop.SolDropUp enabled
	if Enabled then sw77.visible = True
End Sub



'//////////////////////////////////////////////////////////////////////
'// RAMP SWITCHES 
'//////////////////////////////////////////////////////////////////////

Sub SW61_Hit:Controller.Switch(61) = 1:PlaySound "":End Sub
Sub SW61_UnHit:Controller.Switch(61) = 0:End Sub

Sub SW62_Hit:Controller.Switch(62) = 1:PlaySound "":End Sub
Sub SW62_UnHit:Controller.Switch(62) = 0:End Sub

Sub SW63_Hit:Controller.Switch(63) = 1:PlaySound "":End Sub
Sub SW63_UnHit:Controller.Switch(63) = 0:End Sub

Sub SW64_Hit:Controller.Switch(64) = 1:PlaySound "":End Sub
Sub SW64_UnHit:Controller.Switch(64) = 0:End Sub

Sub SW65_Hit:Controller.Switch(65) = 1:PlaySound "fx_metalrolling",0,Vol(ActiveBall)+2,Pan(ActiveBall),0,0,1,0,AudioFade(ActiveBall):End Sub
Sub SW65_UnHit:Controller.Switch(65) = 0:End Sub


'//////////////////////////////////////////////////////////////////////
'// ROLLOVER SWITCHES 
'//////////////////////////////////////////////////////////////////////

Sub SW16_Hit:Controller.Switch(16) = 1:PlaySound "rollover",0,Vol(ActiveBall)+2,Pan(ActiveBall),0,0,1,0,AudioFade(ActiveBall):End Sub
Sub SW16_UnHit:Controller.Switch(16) = 0:End Sub

Sub SW17_Hit:Controller.Switch(17) = 1:PlaySound "rollover",0,Vol(ActiveBall)+2,Pan(ActiveBall),0,0,1,0,AudioFade(ActiveBall):End Sub
Sub SW17_UnHit:Controller.Switch(17) = 0:End Sub

Sub SW26_Hit:Controller.Switch(26) = 1:PlaySound "rollover",0,Vol(ActiveBall)+2,Pan(ActiveBall),0,0,1,0,AudioFade(ActiveBall):End Sub
Sub SW26_UnHit:Controller.Switch(26) = 0:End Sub

Sub SW27_Hit:Controller.Switch(27) = 1:PlaySound "rollover",0,Vol(ActiveBall)+2,Pan(ActiveBall),0,0,1,0,AudioFade(ActiveBall):End Sub
Sub SW27_UnHit:Controller.Switch(27) = 0:End Sub


Sub SW38_Hit:Controller.Switch(38) = 1:PlaySound "rollover",0,Vol(ActiveBall)+1,Pan(ActiveBall),0,0,1,0,AudioFade(ActiveBall):End Sub
Sub SW38_UnHit:Controller.Switch(38) = 0:End Sub

Sub SW48_Hit:Controller.Switch(48) = 1:PlaySound "rollover",0,Vol(ActiveBall)+1,Pan(ActiveBall),0,0,1,0,AudioFade(ActiveBall):End Sub
Sub SW48_UnHit:Controller.Switch(48) = 0:End Sub

Sub SW71_Hit:Controller.Switch(71) = 1:PlaySound "rollover",0,Vol(ActiveBall)+1,Pan(ActiveBall),0,0,1,0,AudioFade(ActiveBall):End Sub
Sub SW71_UnHit:Controller.Switch(71) = 0:End Sub

Sub SW72_Hit:Controller.Switch(72) = 1:PlaySound "rollover",0,Vol(ActiveBall)+1,Pan(ActiveBall),0,0,1,0,AudioFade(ActiveBall):End Sub
Sub SW72_UnHit:Controller.Switch(72) = 0:End Sub

Sub SW73_Hit:Controller.Switch(73) = 1:PlaySound "rollover",0,Vol(ActiveBall)+1,Pan(ActiveBall),0,0,1,0,AudioFade(ActiveBall):End Sub
Sub SW73_UnHit:Controller.Switch(73) = 0:End Sub

Sub SW74_Hit:Controller.Switch(74) = 1:PlaySound "rollover",0,Vol(ActiveBall)+1,Pan(ActiveBall),0,0,1,0,AudioFade(ActiveBall):End Sub
Sub SW74_UnHit:Controller.Switch(74) = 0:End Sub

'//////////////////////////////////////////////////////////////////////
'// SAUCER MOVEMENT 
'//////////////////////////////////////////////////////////////////////
Dim cBall, mMagnet

Sub WobbleMagnet_Init
	 Set mMagnet = new cvpmMagnet
	 With mMagnet
		.InitMagnet WobbleMagnet, 2
		.Size = 100
		.CreateEvents mMagnet
		.MagnetOn = True
	 End With
	Set cBall = ckicker.createball:ckicker.Kick 0,0:mMagnet.addball cball

End Sub

Sub SolUfoShake(Enabled) 
If Enabled Then
BigUfoShake
PlaySound SoundFX("SaucerShake",DOFShaker),0,2,0,0,0,0,1,-.3
End If
End Sub

Sub BigUfoShake
cball.velx = 4 
cball.vely = -18
End Sub

Sub UfoShaker_Timer
Dim a, b, c
a = 90-((ckicker.y - cball.y) / 5)
b = (ckicker.y - cball.y) / 10
c = (cball.x - ckicker.x) / 5




Ufo.rotx = a:Ufo.transz = b:Ufo.rotz = c
'LED1.rotx = a:LED1.transz = b:LED1.rotz = c
'LED2.rotx = a:LED2.transz = b:LED2.rotz = c
'LED3.rotx = a:LED3.transz = b:LED3.rotz = c
LED4.rotx = a:LED4.transz = b:LED4.rotz = c
LED5.rotx = a:LED5.transz = b:LED5.rotz = c
LED6.rotx = a:LED6.transz = b:LED6.rotz = c
LED7.rotx = a:LED7.transz = b:LED7.rotz = c
LED8.rotx = a:LED8.transz = b:LED8.rotz = c
LED9.rotx = a:LED9.transz = b:LED9.rotz = c
LED10.rotx = a:LED10.transz = b:LED10.rotz = c
LED11.rotx = a:LED11.transz = b:LED11.rotz = c
LED12.rotx = a:LED12.transz = b:LED12.rotz = c
LED13.rotx = a:LED13.transz = b:LED13.rotz = c
'LED14.rotx = a:LED14.transz = b:LED14.rotz = c
'LED15.rotx = a:LED15.transz = b:LED15.rotz = c
'LED16.rotx = a:LED16.transz = b:LED16.rotz = c
End Sub

'//////////////////////////////////////////////////////////////////////
'// ALIEN MOVEMENTS 
'//////////////////////////////////////////////////////////////////////

Dim Alien1Pos

Sub Alien1Move (enabled)
if enabled then
Alien1Timer.Enabled=1
PlaySoundAtVol SoundFX("AlienShake1",DOFShaker),Licht14,2
End If
End Sub

Sub Alien1Timer_Timer()	
Select Case Alien1Pos
Case 0: Alien_LF.Z=60:AlienPost1.Z=-60
Case 1: Alien_LF.Z=65:AlienPost1.Z=-55
Case 2: Alien_LF.Z=70:AlienPost1.Z=-50
Case 3: Alien_LF.Z=75:AlienPost1.Z=-45
Case 4: Alien_LF.Z=80:AlienPost1.Z=-40
Case 5: Alien_LF.Z=85:AlienPost1.Z=-35
Case 6: Alien_LF.Z=90:AlienPost1.Z=-30
Case 7: Alien_LF.Z=95:AlienPost1.Z=-25
Case 8: Alien_LF.Z=100:AlienPost1.Z=-20
Case 9: Alien_LF.Z=95:AlienPost1.Z=-25
Case 10: Alien_LF.Z=90:AlienPost1.Z=-30
Case 11: Alien_LF.Z=85:AlienPost1.Z=-35
Case 12: Alien_LF.Z=80:AlienPost1.Z=-40
Case 13: Alien_LF.Z=75:AlienPost1.Z=-45
Case 14: Alien_LF.Z=70:AlienPost1.Z=-50
Case 15: Alien_LF.Z=65:AlienPost1.Z=-55
Case 16: Alien_LF.Z=60:AlienPost1.Z=-60:Alien1Pos=0:Alien1Timer.Enabled=0
End Select

If Alien1Pos=>0 then Alien1Pos=Alien1Pos+1
End Sub


Dim Alien2Pos

Sub Alien2Move (enabled)
if enabled then
Alien2Timer.Enabled=1
PlaySoundAtVol SoundFX("AlienShake2",DOFShaker),f28a,2
End If
End Sub

Sub Alien2Timer_Timer()	
Select Case Alien2Pos
Case 0: Alien_LR.Z=60:AlienPost2.Z=-60
Case 1: Alien_LR.Z=65:AlienPost2.Z=-55
Case 2: Alien_LR.Z=70:AlienPost2.Z=-50
Case 3: Alien_LR.Z=75:AlienPost2.Z=-45
Case 4: Alien_LR.Z=80:AlienPost2.Z=-40
Case 5: Alien_LR.Z=85:AlienPost2.Z=-35
Case 6: Alien_LR.Z=90:AlienPost2.Z=-30
Case 7: Alien_LR.Z=95:AlienPost2.Z=-25
Case 8: Alien_LR.Z=100:AlienPost2.Z=-20
Case 9: Alien_LR.Z=95:AlienPost2.Z=-25
Case 10: Alien_LR.Z=90:AlienPost2.Z=-30
Case 11: Alien_LR.Z=85:AlienPost2.Z=-35
Case 12: Alien_LR.Z=80:AlienPost2.Z=-40
Case 13: Alien_LR.Z=75:AlienPost2.Z=-45
Case 14: Alien_LR.Z=70:AlienPost2.Z=-50
Case 15: Alien_LR.Z=65:AlienPost2.Z=-55:Alien2Pos=0:Alien2Timer.Enabled=0
End Select

If Alien2Pos=>0 then Alien2Pos=Alien2Pos+1
End Sub

Dim Alien3Pos

Sub Alien3Move (enabled)
if enabled then
Alien3Timer.Enabled=1
PlaySoundAtVol SoundFX("AlienShake3",DOFShaker),f28a,2
End If
End Sub

Sub Alien3Timer_Timer()	
Select Case Alien3Pos
Case 0: Alien_RR.Z=60:AlienPost3.Z=-60
Case 1: Alien_RR.Z=65:AlienPost3.Z=-55
Case 2: Alien_RR.Z=70:AlienPost3.Z=-50
Case 3: Alien_RR.Z=75:AlienPost3.Z=-45
Case 4: Alien_RR.Z=80:AlienPost3.Z=-40
Case 5: Alien_RR.Z=85:AlienPost3.Z=-35
Case 6: Alien_RR.Z=90:AlienPost3.Z=-30
Case 7: Alien_RR.Z=95:AlienPost3.Z=-25
Case 8: Alien_RR.Z=100:AlienPost3.Z=-20
Case 9: Alien_RR.Z=95:AlienPost3.Z=-25
Case 10: Alien_RR.Z=90:AlienPost3.Z=-30
Case 11: Alien_RR.Z=85:AlienPost3.Z=-35
Case 12: Alien_RR.Z=80:AlienPost3.Z=-40
Case 13: Alien_RR.Z=75:AlienPost3.Z=-45
Case 14: Alien_RR.Z=70:AlienPost3.Z=-50
Case 15: Alien_RR.Z=65:AlienPost3.Z=-55:Alien3Pos=0:Alien3Timer.Enabled=0
End Select


If Alien3Pos=>0 then Alien3Pos=Alien3Pos+1
End Sub


Dim Alien4Pos

Sub Alien4Move (enabled)
if enabled then
Alien4Timer.Enabled=1
PlaySoundAtVol SoundFX("AlienShake4",DOFShaker),Licht9,2
End If
End Sub

Sub Alien4Timer_Timer()	
Select Case Alien4Pos
Case 0: Alien_RF.Z=60:AlienPost4.Z=-60
Case 1: Alien_RF.Z=65:AlienPost4.Z=-55
Case 2: Alien_RF.Z=70:AlienPost4.Z=-50
Case 3: Alien_RF.Z=75:AlienPost4.Z=-45
Case 4: Alien_RF.Z=80:AlienPost4.Z=-40
Case 5: Alien_RF.Z=85:AlienPost4.Z=-35
Case 6: Alien_RF.Z=90:AlienPost4.Z=-30
Case 7: Alien_RF.Z=95:AlienPost4.Z=-25
Case 8: Alien_RF.Z=100:AlienPost4.Z=-20
Case 9: Alien_RF.Z=95:AlienPost4.Z=-25
Case 10: Alien_RF.Z=90:AlienPost4.Z=-30
Case 11: Alien_RF.Z=85:AlienPost4.Z=-35
Case 12: Alien_RF.Z=80:AlienPost4.Z=-40
Case 13: Alien_RF.Z=75:AlienPost4.Z=-45
Case 14: Alien_RF.Z=70:AlienPost4.Z=-50
Case 15: Alien_RF.Z=65:AlienPost4.Z=-55:Alien4Pos=0:Alien4Timer.Enabled=0
End Select

If Alien4Pos=>0 then Alien4Pos=Alien4Pos+1
End Sub

'//////////////////////////////////////////////////////////////////////
'// TARGETBANK MOVEMENT 
'//////////////////////////////////////////////////////////////////////

Dim TBPos, TBDown

Sub TBMove (enabled)
	TBTimer.Enabled = Enabled
	if enabled then PlaySoundAt SoundFX("TargetBank",DOFContactors),DPTrigger
End Sub

Sub TBTimer_Timer()	
Select Case TBPos
Case 0: CENTREBANK.Z=-20:SW45P.Z=-20:SW46P.Z=-20:SW47P.Z=-20:TBPos=0:Controller.Switch(66) = 0:Controller.Switch(67) = 1::SW45.isdropped=0:SW46.isdropped=0:SW47.isdropped=0:DPWall.isdropped=0:DPWall1.isdropped=1:DPRAMP.collidable=1
Case 1: CENTREBANK.Z=-22:SW45P.Z=-22:SW46P.Z=-22:SW47P.Z=-22:Controller.Switch(67) = 0
Case 2: CENTREBANK.Z=-24:SW45P.Z=-24:SW46P.Z=-24:SW47P.Z=-24
Case 3: CENTREBANK.Z=-26:SW45P.Z=-26:SW46P.Z=-26:SW47P.Z=-26
Case 4: CENTREBANK.Z=-28:SW45P.Z=-28:SW46P.Z=-28:SW47P.Z=-28
Case 5: CENTREBANK.Z=-30:SW45P.Z=-30:SW46P.Z=-30:SW47P.Z=-30
Case 6: CENTREBANK.Z=-32:SW45P.Z=-32:SW46P.Z=-32:SW47P.Z=-32
Case 7: CENTREBANK.Z=-34:SW45P.Z=-34:SW46P.Z=-34:SW47P.Z=-34
Case 8: CENTREBANK.Z=-36:SW45P.Z=-36:SW46P.Z=-36:SW47P.Z=-36
Case 9: CENTREBANK.Z=-38:SW45P.Z=-38:SW46P.Z=-38:SW47P.Z=-38
Case 10: CENTREBANK.Z=-40:SW45P.Z=-40:SW46P.Z=-40:SW47P.Z=-40
Case 11: CENTREBANK.Z=-42:SW45P.Z=-42:SW46P.Z=-42:SW47P.Z=-42
Case 12: CENTREBANK.Z=-44:SW45P.Z=-44:SW46P.Z=-44:SW47P.Z=-44:
Case 13: CENTREBANK.Z=-46:SW45P.Z=-46:SW46P.Z=-46:SW47P.Z=-46:
Case 14: CENTREBANK.Z=-48:SW45P.Z=-48:SW46P.Z=-48:SW47P.Z=-48
Case 15: CENTREBANK.Z=-50:SW45P.Z=-50:SW46P.Z=-50:SW47P.Z=-50
Case 16: CENTREBANK.Z=-52:SW45P.Z=-52:SW46P.Z=-52:SW47P.Z=-52
Case 17: CENTREBANK.Z=-54:SW45P.Z=-54:SW46P.Z=-54:SW47P.Z=-54
Case 18: CENTREBANK.Z=-56:SW45P.Z=-56:SW46P.Z=-56:SW47P.Z=-56
Case 19: CENTREBANK.Z=-58:SW45P.Z=-58:SW46P.Z=-58:SW47P.Z=-58
Case 20: CENTREBANK.Z=-60:SW45P.Z=-60:SW46P.Z=-60:SW47P.Z=-60
Case 21: CENTREBANK.Z=-62:SW45P.Z=-62:SW46P.Z=-62:SW47P.Z=-62
Case 22: CENTREBANK.Z=-64:SW45P.Z=-64:SW46P.Z=-64:SW47P.Z=-64
Case 23: CENTREBANK.Z=-66:SW45P.Z=-66:SW46P.Z=-66:SW47P.Z=-66
Case 24: CENTREBANK.Z=-68:SW45P.Z=-68:SW46P.Z=-68:SW47P.Z=-68
Case 25: CENTREBANK.Z=-70:SW45P.Z=-70:SW46P.Z=-70:SW47P.Z=-70
Case 26: CENTREBANK.Z=-72:SW45P.Z=-72:SW46P.Z=-72:SW47P.Z=-72
Case 27: CENTREBANK.Z=-74:SW45P.Z=-74:SW46P.Z=-74:SW47P.Z=-74
Case 28: CENTREBANK.Z=-76:SW45P.Z=-76:SW46P.Z=-76:SW47P.Z=-76:SW45.isdropped=1:SW46.isdropped=1:SW47.isdropped=1:DPWALL.isdropped=1:DPRAMP.collidable=0:Controller.Switch(66) = 0
Case 29: Controller.Switch(66) = 1:Controller.Switch(67) = 0
End Select

If TBDown=0 then TBPos=TBPos+1 
If TBDown=1 then TBPos=TBPos-1
if TBPos < -2 then TBDown = 0 else if TBPos > 31 then TBDown = 1
End Sub

'//////////////////////////////////////////////////////////////////////
'// DIRTY POOL HANDLE 
'//////////////////////////////////////////////////////////////////////

Dim DPBall, Ballresting

DPBall = ""
Ballresting = False

Sub DPTrigger_Hit
Ballresting = True
Set DPBall = ActiveBall
End Sub

Sub DPTrigger_UnHit
Ballresting = False
DPRAMP.collidable=0
End Sub

'//////////////////////////////////////////////////////////////////////
'// DIVERTER MOVEMENT 
'//////////////////////////////////////////////////////////////////////

Dim DivPos, DivClosed, DiverterDir

Sub DivMove(Enabled)
If Enabled Then
PlaySound SoundFX("Diverter",DOFContactors)
DiverterDir = 1
Else
DiverterDir = -1
End If

DivTimer.Enabled = 0
If DivPos <1 Then DivPos = 1
If DivPos > 8 Then DivPos = 8

DivTimer.Enabled = 1
End Sub

Sub DivTimer_Timer()
Select Case DivPos
Case 0:Diverter.ObjRotZ=-15:Diverter9.IsDropped = 0:Diverter8.IsDropped = 1:DivTimer.Enabled = 0
Case 1:Diverter.ObjRotZ=-8:Diverter8.IsDropped = 0:Diverter9.IsDropped = 1:Diverter7.IsDropped = 1
Case 2:Diverter.ObjRotZ=-3:Diverter7.IsDropped = 0:Diverter8.IsDropped = 1:Diverter6.IsDropped = 1
Case 3:Diverter.ObjRotZ=2:Diverter6.IsDropped = 0:Diverter7.IsDropped = 1:Diverter5.IsDropped = 1
Case 4:Diverter.ObjRotZ=7:Diverter5.IsDropped = 0:Diverter6.IsDropped = 1:Diverter4.IsDropped = 1
Case 5:Diverter.ObjRotZ=12:Diverter4.IsDropped = 0:Diverter5.IsDropped = 1:Diverter3.IsDropped = 1
Case 6:Diverter.ObjRotZ=17:Diverter3.IsDropped = 0:Diverter4.IsDropped = 1:Diverter2.IsDropped = 1
Case 7:Diverter.ObjRotZ=22:Diverter2.IsDropped = 0:Diverter3.IsDropped = 1:Diverter1.IsDropped = 1
Case 8:Diverter.ObjRotZ=27:Diverter1.IsDropped = 0:Diverter2.isdropped = 0
Case 9:DivTimer.Enabled = 0
End Select

DivPos = DivPos + DiverterDir
End Sub

'//////////////////////////////////////////////////////////////////////
'// OBJECT SOUND EFFECTS 
'//////////////////////////////////////////////////////////////////////

Sub LeftPost_Hit(): PlaySoundAt "RubberHit",ActiveBall: End Sub

Sub LeftSlingRubber_Hit(): PlaySoundAt "RubberHit",ActiveBall:End Sub

Sub RightPost_Hit(): PlaySoundAt "RubberHit",ActiveBall: End Sub

Sub RightSlingRubber_Hit(): PlaySoundAt "RubberHit",ActiveBall:End Sub

Sub Rubber3_Hit(): PlaySoundAt "RubberHit3",ActiveBall: End Sub

Sub Rubber6_Hit(): PlaySoundAt "RubberHit3",ActiveBall: End Sub 

Sub Wall14_Hit(): PlaySoundAt "Objecthit",ActiveBall:End Sub

Sub Wall31_Hit(): PlaySoundAt "Objecthit",ActiveBall:End Sub

Sub Wall29_Hit(): PlaySoundAt "Objecthit",ActiveBall:End Sub

Sub Wall34_Hit(): PlaySoundAt "Objecthit",ActiveBall:End Sub




Sub LeftFlipper_Collide(parm):
Ballspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
If Ballspeed >2 then Playsound "RubberHit2" Else Playsound "RubberHitLow":End If:End Sub

Sub RightFlipper_Collide(parm):
Ballspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
If Ballspeed >2 then Playsound "RubberHit2" Else Playsound "RubberHitLow":End If:End Sub

Sub RDrop_Hit()
ActiveBall.VelZ = -2
ActiveBall.VelY = 0
ActiveBall.VelX = 0
StopSound "fx_metalrolling"
PlaySound "Balldrop",0,1,.4,0,0,0,1,.8
End Sub

Sub LDrop_Hit()
ActiveBall.VelZ = -2
ActiveBall.VelY = 0
ActiveBall.VelX = 0
StopSound "fx_metalrolling"
PlaySound "Balldrop",0,1,-.4,0,0,0,1,.8
End Sub

Sub LRoll_Hit()
PlaySound "fx_metalrolling",0,Vol(ActiveBall)+2,Pan(ActiveBall),0,0,1,0,AudioFade(ActiveBall)
End Sub

'//////////////////////////////////////////////////////////////////////
'// BALL SOUND FUNCTIONS 
'//////////////////////////////////////////////////////////////////////

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
Vol = Csng(BallVel(ball) ^2 / 1000)
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "AFM" is the name of the table
Dim tmp
tmp = ball.x * 2 / AFM.width-1
If tmp > 0 Then
Pan = Csng(tmp ^10)
Else
Pan = Csng(-((- tmp) ^10) )
End If
End Function

function AudioFade(ball)
	Dim tmp
    tmp = ball.y * 2 / AFM.height-1
    If tmp > 0 Then
		AudioFade = Csng(tmp ^10)
    Else
        AudioFade = Csng(-((- tmp) ^10) )
    End If
End Function

Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
Pitch = BallVel(ball) * 20
End Function

Function BallVel(ball) 'Calculates the ball speed
BallVel = INT(SQR((ball.VelX ^2) + (ball.VelY ^2) ) )
End Function

'//////////////////////////////////////////////////////////////////////
'// JP VPX ROLLING SOUNDS 
'//////////////////////////////////////////////////////////////////////

Const tnob = 5 ' total number of balls in this table is 4, but always use a higher number here because of the timing
ReDim rolling(tnob)
InitRolling

Sub InitRolling
Dim i
For i = 0 to tnob
rolling(i) = False
Next
End Sub

Sub RollingUpdate()
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
						PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) )*.6, Pan(BOT(b) ), 0, Pitch(BOT(b) )+50000, 1, 0, AudioFade(BOT(b) )
				End If
		Else
            If rolling(b) = True Then
                StopSound("fx_ballrolling" & b)
                rolling(b) = False
            End If
        End If
    Next
End Sub

'//////////////////////////////////////////////////////////////////////
'// BALL COLLISION SOUNDS 
'//////////////////////////////////////////////////////////////////////

Sub OnBallBallCollision(ball1, ball2, velocity)
PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 200, Pan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub

'//////////////////////////////////////////////////////////////////////
'// Positional Sound Playback Functions by DJRobX
'//////////////////////////////////////////////////////////////////////

'Set position as table object (Use object or light but NOT wall) and Vol to 1

Sub PlaySoundAt(sound, tableobj)
		PlaySound sound, 1, 1, Pan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub


'Set all as per ball position & speed.

Sub PlaySoundAtBall(sound)
		PlaySound sound, 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 1, AudioFade(ActiveBall)
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


Dim NextOrbitHit:NextOrbitHit = 0 

Sub WireRampBumps_Hit(idx)
	if BallVel(ActiveBall) > .3 and Timer > NextOrbitHit then
		RandomBump3 .5, Pitch(ActiveBall)+5
		' Schedule the next possible sound time.  This prevents it from rapid-firing noises too much. 
		' Lowering these numbers allow more closely-spaced clunks.
		NextOrbitHit = Timer + .2 + (Rnd * .2)
	end if 
End Sub

Sub PlasticRampBumps_Hit(idx)
	if BallVel(ActiveBall) > .3 and Timer > NextOrbitHit then
		RandomBump 2, Pitch(ActiveBall)
		' Schedule the next possible sound time.  This prevents it from rapid-firing noises too much. 
		' Lowering these numbers allow more closely-spaced clunks.
		NextOrbitHit = Timer + .1 + (Rnd * .2)
	end if 
End Sub


Sub MetalGuideBumps_Hit(idx)
	if BallVel(ActiveBall) > .3 and Timer > NextOrbitHit then
		RandomBump2 2, Pitch(ActiveBall)
		' Schedule the next possible sound time.  This prevents it from rapid-firing noises too much. 
		' Lowering these numbers allow more closely-spaced clunks.
		NextOrbitHit = Timer + .2 + (Rnd * .2)
	end if 
End Sub

Sub MetalWallBumps_Hit(idx)
	if BallVel(ActiveBall) > .3 and Timer > NextOrbitHit then
		RandomBump 2, 20000 'Increased pitch to simulate metal wall
		' Schedule the next possible sound time.  This prevents it from rapid-firing noises too much. 
		' Lowering these numbers allow more closely-spaced clunks.
		NextOrbitHit = Timer + .2 + (Rnd * .2)
	end if 
End Sub


'' Requires rampbump1 to 7 in Sound Manager
Sub RandomBump(voladj, freq)
	dim BumpSnd:BumpSnd= "rampbump" & CStr(Int(Rnd*7)+1)
		PlaySound BumpSnd, 0, Vol(ActiveBall)*voladj, Pan(ActiveBall), 0, freq, 0, 1, AudioFade(ActiveBall)
End Sub

' Requires metalguidebump1 to 2 in Sound Manager
Sub RandomBump2(voladj, freq)
	dim BumpSnd:BumpSnd= "metalguidebump" & CStr(Int(Rnd*2)+1)
		PlaySound BumpSnd, 0, Vol(ActiveBall)*voladj, Pan(ActiveBall), 0, freq, 0, 1, AudioFade(ActiveBall)
End Sub

' Requires WireRampBump1 to 5 in Sound Manager
Sub RandomBump3(voladj, freq)
	dim BumpSnd:BumpSnd= "WireRampBump" & CStr(Int(Rnd*5)+1)
		PlaySound BumpSnd, 0, Vol(ActiveBall)*voladj, Pan(ActiveBall), 0, freq, 0, 1, AudioFade(ActiveBall)
End Sub



' Stop Bump Sounds - Place triggers for these at the end of the ramps, to ensure that no bump sounds are played after the ball leaves the ramp.

Sub BumpSTOP1_Hit ()
dim i:for i=1 to 4:StopSound "WireRampBump" & i:next
NextOrbitHit = Timer + 1
End Sub

Sub BumpSTOP2_Hit ()
dim i:for i=1 to 4:StopSound "WireRampBump" & i:next
NextOrbitHit = Timer + 1
End Sub




'//////////////////////////////////////////////////////////////////////
'// GI CONTROL 
'//////////////////////////////////////////////////////////////////////


Sub UpdateGI(no, step)
Dim gistep, ii, a
gistep = step / 8
Select Case no
Case 0
if Round(8*gistep) < 1.25 then PLASTICS_Bottom.visible = 0 else PLASTICS_Bottom.visible = 1 end If
PLASTICS_Bottom.material = "GIFade" & Round(8*gistep)
STARPOSTS.material = "GIFade" & Round(8*gistep)

If ApronGIProper = 1 Then
	APRON.material = "GIFade" & Round(8*gistep)
Else
	APRON.material = "GIFadeNoOpacity" & Round(8*gistep)
End If

batleft.material = "FlipperPlastic" & Round(8*gistep)
batright.material = "FlipperPlastic" & Round(8*gistep)
For each ii in GIBottom
ii.IntensityScale = gistep
currGIBot = 8*gistep
Next
Case 1
if Round(8*gistep) < 1.25 then PLASTICS_MidL.visible = 0 else PLASTICS_MidL.visible = 1 end If
if Round(8*gistep) < 1.25 then PLASTICS_MidR.visible = 0 else PLASTICS_MidR.visible = 1 end If
PLASTICS_MidL.material = "GIFade" & Round(8*gistep)
PLASTICS_MidR.material = "GIFade" & Round(8*gistep)
SIDEWALL.material = "GIFade" & Round(8*gistep)
if Round(8*gistep) < 1.25 then SIDEWALL.visible = 0 else SIDEWALL.visible = 1 end If
if Round(8*gistep) >= 7 then SIDEWALL_GIOFF.visible = 0 else SIDEWALL_GIOFF.visible = 1 end If
LEFTPOPPERSCOOP.material = "GIFade" & Round(8*gistep)
SW56P.material = "GIFade" & Round(8*gistep)
SW57P.material = "GIFade" & Round(8*gistep)
SW58P.material = "GIFade" & Round(8*gistep)
SW42P.material = "GIFade" & Round(8*gistep)
SW41P.material = "GIFade" & Round(8*gistep)
SAUCERSML.material = "GIFade" &  Round(8*gistep)
STEELWALLS.material = "GIFadeNoOpacity" & Round(8*gistep)
Diverter.material = "GIFadeNoOpacity" & Round(8*gistep)
For each ii in GIMiddle
ii.IntensityScale = gistep
currGIMid = 8*gistep
Next
Case 2
if Round(8*gistep) < 1.25 then PLASTICS_Top.visible = 0 else PLASTICS_Top.visible = 1 end If
PLASTICS_Top.material = "GIFade" & Round(8*gistep)
CLEARAMPS.material = "ClearRamp1" & Round(8*gistep)

if Round(8*gistep) < 1.1 then 
	CLEARAMPS.image = "CLEARAMPS_A"
	Alien_LF.image = "ALIEN_LF_A"
	Alien_LR.image = "ALIEN_LR_A"
	Alien_RF.image = "ALIEN_RF_A"
	Alien_RR.image = "ALIEN_RR_A"
	Ufo.image = "UFO_GIOff"
		else 
	CLEARAMPS.image = "CLEARAMPS_B"
	Alien_LF.image = "ALIEN_LF_A"
	Alien_LR.image = "ALIEN_LR_A"
	Alien_RF.image = "ALIEN_RF_A"
	Alien_RR.image = "ALIEN_RR_A"
	Ufo.image = "UFO_RG"
end If

'REARWALL.material = "GIFadeNoOpacity" & Round(8*gistep)
REDPLASTICS.material = "GIFade" & Round(8*gistep)
SW43P.material = "GIFade" & Round(8*gistep)
SW44P.material = "GIFade" & Round(8*gistep)
For each ii in GITop
ii.IntensityScale = gistep
currGITop = 8*gistep
Next
End Select
End Sub


'//////////////////////////////////////////////////////////////////////
'// REALTIME UPDATES 
'//////////////////////////////////////////////////////////////////////

Set MotorCallback = GetRef("RealTimeUpdates")

Sub RealTimeUpdates
RollingUpdate
End Sub

'//////////////////////////////////////////////////////////////////////
'// JP FADING LIGHT SYSTEM 
'//////////////////////////////////////////////////////////////////////

Dim LampState(200), FadingLevel(200), ModLevel(200)
Dim FlashSpeedUp(200), FlashSpeedDown(200), FlashMin(200), FlashMax(200), FlashLevel(200)

InitLamps()             ' turn off the lights and flashers and reset them to the default parameters
LampTimer.Interval = -1 'lamp fading speed
LampTimer.Enabled = 1



' Lamp & Flasher Timers


'//PRELOADING LIGHTS ON  START HERE////////////////////////////////////

Dim GIInit: GIInit=90

Sub LampTimer_Timer()

If PreloadMe = 1 Then
	if GIInit > 0 Then
		GIInit = GIInit -1
		if GIInit <= 90 and GIInit > 82 then
			UpdateGI 0, GIInit-82
			UpdateGI 1, GIInit-82
			UpdateGI 2, GIInit-82
		Else
			if GIInit < 40 then 
				UpdateGI 0, 8:UpdateGI 1, 8:UpdateGI 2, 8
			elseif GIInit >= 40 and GIInit <=82 then 
				UpdateGI 0, 0:UpdateGI 1, 0:UpdateGI 2, 0
			end if 
		end if 
		if (GIInit mod 40) > 1 then SetLamp 90 + (GIInit mod 40) +1, 1
		SetLamp 90 + (GIInit mod 40) + 2, 0
	end if 
End If


Dim chgLamp, num, chg, ii
chgLamp = Controller.ChangedLamps
If Not IsEmpty(chgLamp) Then
For ii = 0 To UBound(chgLamp)
LampState(chgLamp(ii, 0) ) = chgLamp(ii, 1)       'keep the real state in an array
FadingLevel(chgLamp(ii, 0) ) = chgLamp(ii, 1) + 4 'actual fading step
Next
End If
UpdateLamps
End Sub

'//////////////////////////////////////////////////////////////////////
'// INSERT LIGHTS 
'//////////////////////////////////////////////////////////////////////

Sub UpdateLamps

NFadeLm 11, l11, l11b
NFadeLm 12, l12, l12b
NFadeLm 13, l13, l13b
NFadeLm 14, l14, l14b
NFadeLmm 15, l15l, l15r, l15a
NFadeLm 16, l16, l16b
NFadeLm 17, l17, l17b
NFadeLmm 18, l18, l18b, l18a
NFadeLm 21, l21, l21b
NFadeLm 22, l22, l22b
NFadeLm 23, l23, l23b
NFadeLm 24, l24, l24b
NFadeLmm 25, l25, l25b, l25a
NFadeLm 26, l26, l26b
NFadeLm 27, l27, l27b
NFadeLm 28, l28, l28b
NFadeLm 31, l31, l31b
NFadeLm 32, l32, l32b
NFadeLm 33, l33, l33b
NFadeLm 34, l34, l34b
NFadeLm 35, l35, l35b
NFadeLm 36, l36, l36b
NFadeLm 37, l37, l37b
NFadeLm 38, l38, l38b
NFadeLmm 41, l41, l41b, l41a
NFadeLmm 42, l42, l42b, l42a
NFadeLmm 43, l43, l43b, l43a
NFadeLmm 44, l44, l44b, l44a
NFadeLmm 45, l45, l45b, l45a
NFadeLmm 46, l46, l46b, l46a
NFadeLmm 47, l47, l47b, l47a
NFadeLmm 48, l48, l48b, l48a
NFadeLm 51, l51, l51b
NFadeLm 52, l52, l52b
NFadeLm 53, l53, l53b
NFadeLm 54, l54, l54b
NFadeLm 55, l55, l55b
NFadeLm 56, l56, l56b
NFadeLm 57, l57, l57b
NFadeLm 58, l58, l58b
NFadeLm 61, l61, l61b
NFadeLm 62, l62, l62b
NFadeLm 63, l63, l63b
NFadeLm 64, l64, l64b
NFadeLm 65, l65, l65b
NFadeLm 66, l66, l66b
NFadeLm 67, l67, l67b
NFadeLmm 68, l68, l68b, l68a
NFadeLmm 71, l71, l71b, l71a
NFadeLm 72, l72, l72b
NFadeLm 73, l73, l73b
NFadeLm 74, l74, l74b
NFadeLmm 75, l75, l75b, l75a
NFadeLm 76, l76, l76b
NFadeLm 77, l77, l77b
NFadeLm 78, l78, l78b
NFadeLmm 81, l81, l81b, l81a
NFadeLmm 82, l82, l82b, l82a
NFadeLmm 83, l83, l83b, l83a
NFadeLmm 84, l84, l84b, l84a
NFadeLmm 85, l85, l85b, l85a

'86 - Launch button LED
'88 - Start button LED

'NFadeL 86, l86
'NFadeL 88, l88

'//////////////////////////////////////////////////////////////////////
'// FLASHER 
'//////////////////////////////////////////////////////////////////////

NFadeLmf 117, F17dr, F17flare
NFadeLmf 118, F18dr, F18flare
NFadeLmf 119, F19dr, F19flare
NFadeLmf 125, f25dr, F25flare
NFadeLmf 126, F26dr, F26flare
NFadeLmf 127, f27dr, F27flare





'NFadeLmm 117, F17, F17b, F17c
'NFadeLmm 118, F18, F18b, F18c
'NFadeLmm 119, F19, F19b, F19c


NFadeLmf 120, f20, f20a
NFadeLm 121, F21, F21a
NFadeLmf 123, f23, f23c
'NFadeLmm 125, F25, F25b, F25c
'NFadeLmm 126, F26, F26b, F26c
'NFadeLmm 127, F27, F27b, F27c
NFadeLmf 128, f28, f28a




'//////////////////////////////////////////////////////////////////////
'// SAUCER LED 
'//////////////////////////////////////////////////////////////////////


FadeObj 94, Led4, "UFOLED_A", "UFOLED_B", "UFOLED_C", "UFOLED_D"
FadeObj 95, Led5, "UFOLED_A", "UFOLED_B", "UFOLED_C", "UFOLED_D"
FadeObj 96, Led6, "UFOLED_A", "UFOLED_B", "UFOLED_C", "UFOLED_D"
FadeObj 97, Led7, "UFOLED_A", "UFOLED_B", "UFOLED_C", "UFOLED_D"
FadeObj 98, Led8, "UFOLED_A", "UFOLED_B", "UFOLED_C", "UFOLED_D"
FadeObj 101, Led9, "UFOLED_A", "UFOLED_B", "UFOLED_C", "UFOLED_D"
FadeObj 102, Led10, "UFOLED_A", "UFOLED_B", "UFOLED_C", "UFOLED_D"
FadeObj 103, Led11, "UFOLED_A", "UFOLED_B", "UFOLED_C", "UFOLED_D"
FadeObj 104, Led12, "UFOLED_A", "UFOLED_B", "UFOLED_C", "UFOLED_D"
FadeObj 105, Led13, "UFOLED_A", "UFOLED_B", "UFOLED_C", "UFOLED_D"


'// These LEDS are round back so not visibile - Off for performance ///////
'FadeObj 91, Led1, "UFOLED_A", "UFOLED_B", "UFOLED_C", "UFOLED_D"
'FadeObj 92, Led2, "UFOLED_A", "UFOLED_B", "UFOLED_C", "UFOLED_D"
'FadeObj 93, Led3, "UFOLED_A", "UFOLED_B", "UFOLED_C", "UFOLED_D"

'FadeObj 106, Led14, "UFOLED_A", "UFOLED_B", "UFOLED_C", "UFOLED_D"
'FadeObj 107, Led15, "UFOLED_A", "UFOLED_B", "UFOLED_C", "UFOLED_D"
'FadeObj 108, Led16, "UFOLED_A", "UFOLED_B", "UFOLED_C", "UFOLED_D"




'///////////////////////////////////////////////////////////////////////////
'// FLASHERS - 20 = LeftFlasher 23 = Saucer Flasher 28 = Right Flasher 39 = Strobe
'///////////////////////////////////////////////////////////////////////////


If UFOFlashers =1 Then
	NFadeObjm 120, Ufo, "UFO_RF", "UFO_RG"
	NFadeObjm 123, Ufo, "UFO_Flasher_D", "UFO_RG"
	NFadeObjm 128, Ufo, "UFO_LF", "UFO_RG"
End If

If ChromeRailFlashers = 1 Then

' FLASHERS / STROBE - switchback to GI off texture
	NFadeObjm 120, CHROMERAILS, "CHROMERAILS_D", "CHROMERAILS_A"
	NFadeObjm 123, CHROMERAILS, "CHROMERAILS_C", "CHROMERAILS_A"
	NFadeObjm 128, CHROMERAILS, "CHROMERAILS_E", "CHROMERAILS_A"

' REDFLASHERS - switchback to GI on texture
	NFadeObjm 125, CHROMERAILS, "CHROMERAILS_F", "CHROMERAILS_B"
	NFadeObjm 126, CHROMERAILS, "CHROMERAILS_F", "CHROMERAILS_B"
	NFadeObjm 127, CHROMERAILS, "CHROMERAILS_F", "CHROMERAILS_B"
	NFadeObjm 117, CHROMERAILS, "CHROMERAILS_F", "CHROMERAILS_B"
	NFadeObjm 119, CHROMERAILS, "CHROMERAILS_F", "CHROMERAILS_B"
	NFadeObjm 118, CHROMERAILS, "CHROMERAILS_F", "CHROMERAILS_B"


	NFadeObjm 120, LEFTPOPPERSCOOP_GIOFF, "LEFTPOPPERSCOOP_C", "LEFTPOPPERSCOOP_A"
	NFadeObjm 123, LEFTPOPPERSCOOP_GIOFF, "LEFTPOPPERSCOOP_C", "LEFTPOPPERSCOOP_A"
	NFadeObjm 128, LEFTPOPPERSCOOP_GIOFF, "LEFTPOPPERSCOOP_C", "LEFTPOPPERSCOOP_A"
	NFadeObjm 125, LEFTPOPPERSCOOP_GIOFF, "LEFTPOPPERSCOOP_F", "LEFTPOPPERSCOOP_A"
	NFadeObjm 126, LEFTPOPPERSCOOP_GIOFF, "LEFTPOPPERSCOOP_F", "LEFTPOPPERSCOOP_A"
	NFadeObjm 127, LEFTPOPPERSCOOP_GIOFF, "LEFTPOPPERSCOOP_F", "LEFTPOPPERSCOOP_A"
	NFadeObjm 117, LEFTPOPPERSCOOP_GIOFF, "LEFTPOPPERSCOOP_F", "LEFTPOPPERSCOOP_A"
	NFadeObjm 119, LEFTPOPPERSCOOP_GIOFF, "LEFTPOPPERSCOOP_F", "LEFTPOPPERSCOOP_A"
	NFadeObjm 118, LEFTPOPPERSCOOP_GIOFF, "LEFTPOPPERSCOOP_F", "LEFTPOPPERSCOOP_A"

	NFadeObjm 120, METALBRUSHED, "METALBRUSHED_C", "METALBRUSHED_B"
	NFadeObjm 123, METALBRUSHED, "METALBRUSHED_C", "METALBRUSHED_B"
	NFadeObjm 128, METALBRUSHED, "METALBRUSHED_C", "METALBRUSHED_B"
	NFadeObjm 125, METALBRUSHED, "METALBRUSHED_F", "METALBRUSHED_B"
	NFadeObjm 126, METALBRUSHED, "METALBRUSHED_F", "METALBRUSHED_B"
	NFadeObjm 127, METALBRUSHED, "METALBRUSHED_F", "METALBRUSHED_B"
	NFadeObjm 117, METALBRUSHED, "METALBRUSHED_F", "METALBRUSHED_B"
	NFadeObjm 119, METALBRUSHED, "METALBRUSHED_F", "METALBRUSHED_B"
	NFadeObjm 118, METALBRUSHED, "METALBRUSHED_F", "METALBRUSHED_B"

	NFadeObjm 120, BLACK_PAINTED, "BLACK_PAINTED_C", "BLACK_PAINTED_A"
	NFadeObjm 123, BLACK_PAINTED, "BLACK_PAINTED_C", "BLACK_PAINTED_A"
	NFadeObjm 128, BLACK_PAINTED, "BLACK_PAINTED_C", "BLACK_PAINTED_A"
	NFadeObjm 125, BLACK_PAINTED, "BLACK_PAINTED_F", "BLACK_PAINTED_A"
	NFadeObjm 126, BLACK_PAINTED, "BLACK_PAINTED_F", "BLACK_PAINTED_A"
	NFadeObjm 127, BLACK_PAINTED, "BLACK_PAINTED_F", "BLACK_PAINTED_A"
	NFadeObjm 117, BLACK_PAINTED, "BLACK_PAINTED_F", "BLACK_PAINTED_A"
	NFadeObjm 119, BLACK_PAINTED, "BLACK_PAINTED_F", "BLACK_PAINTED_A"
	NFadeObjm 118, BLACK_PAINTED, "BLACK_PAINTED_F", "BLACK_PAINTED_A"


End If

If PlasticsTopFlasher = 1 Then
	NFadeObjm 123, PLASTICS_Top, "PLASTICS_C", "PLASTICS_B"
	NFadeObjm 123, PLASTICS_Top_GIOFF, "PLASTICS_C", "PLASTICS_A"
	NFadeObjm 123, REDPLASTICS_GIOFF, "RED_PLASTICS_FLASHERS", "REDPLASTICS_A"
	

If currGITop <4 Then
NFadeObjm 123, CLEARAMPS, "CLEARAMPS_C", "CLEARAMPS_A"
Else
NFadeObjm 123, CLEARAMPS, "CLEARAMPS_C", "CLEARAMPS_B"
End If

End If

If AlienFlashers = 1 Then
	NFadeObjm 120, Alien_LF, "ALIEN_LF_D", "ALIEN_LF_A"
	NFadeObjm 120, Alien_RF, "ALIEN_RF_D", "ALIEN_RF_A"
	NFadeObjm 120, Alien_RR, "ALIEN_RR_D", "ALIEN_RR_A"
	NFadeObjm 120, Alien_LR, "ALIEN_LR_D", "ALIEN_LR_A"

	NFadeObjm 123, Alien_LF, "ALIEN_LF_C", "ALIEN_LF_A"
	NFadeObjm 123, Alien_RF, "ALIEN_RF_C", "ALIEN_RF_A"
	NFadeObjm 123, Alien_RR, "ALIEN_RR_C", "ALIEN_RR_A"
	NFadeObjm 123, Alien_LR, "ALIEN_LR_C", "ALIEN_LR_A"

	NFadeObjm 128, Alien_LF, "ALIEN_LF_E", "ALIEN_LF_A"
	NFadeObjm 128, Alien_RF, "ALIEN_RF_E", "ALIEN_RF_A"
	NFadeObjm 128, Alien_RR, "ALIEN_RR_E", "ALIEN_RR_A"
	NFadeObjm 128, Alien_LR, "ALIEN_LR_E", "ALIEN_LR_A"
End If

If SideWallFlashers = 1 Then
	NFadeObjmm 120, SIDEWALL, SIDEWALL_GIOFF, "SIDEWALL_DF", "SIDEWALL_B", "SIDEWALL_A", "SIDEWALL_D"
	NFadeObjmm 123, SIDEWALL, SIDEWALL_GIOFF, "SIDEWALL_CF", "SIDEWALL_B", "SIDEWALL_A", "SIDEWALL_C"
	NFadeObjmm 128, SIDEWALL, SIDEWALL_GIOFF, "SIDEWALL_EF", "SIDEWALL_B", "SIDEWALL_A", "SIDEWALL_E"
	NFadeObjmm 117, SIDEWALL, SIDEWALL_GIOFF, "SIDEWALL_FF", "SIDEWALL_B", "SIDEWALL_A", "SIDEWALL_F"
	NFadeObjmm 118, SIDEWALL, SIDEWALL_GIOFF, "SIDEWALL_FF", "SIDEWALL_B", "SIDEWALL_A", "SIDEWALL_F"
	NFadeObjmm 119, SIDEWALL, SIDEWALL_GIOFF, "SIDEWALL_FF", "SIDEWALL_B", "SIDEWALL_A", "SIDEWALL_F"
	NFadeObjmm 125, SIDEWALL, SIDEWALL_GIOFF, "SIDEWALL_FF", "SIDEWALL_B", "SIDEWALL_A", "SIDEWALL_F"
	NFadeObjmm 126, SIDEWALL, SIDEWALL_GIOFF, "SIDEWALL_FF", "SIDEWALL_B", "SIDEWALL_A", "SIDEWALL_F"
	NFadeObjmm 127, SIDEWALL, SIDEWALL_GIOFF, "SIDEWALL_FF", "SIDEWALL_B", "SIDEWALL_A", "SIDEWALL_F"
End If







End Sub

'//////////////////////////////////////////////////////////////////////
'// LAMP SUBS 
'//////////////////////////////////////////////////////////////////////

Sub InitLamps()
Dim x
For x = 0 to 200
LampState(x) = 0         ' current light state, independent of the fading level. 0 is off and 1 is on
FadingLevel(x) = 4       ' used to track the fading state
ModLevel(x) = 0
FlashSpeedUp(x) = 0.2    ' faster speed when turning on the flasher
FlashSpeedDown(x) = 0.01  ' slower speed when turning off the flasher
FlashMax(x) = 1          ' the maximum value when on, usually 1
FlashMin(x) = 0          ' the minimum value when off, usually 0
FlashLevel(x) = 0        ' the intensity of the flashers, usually from 0 to 1
Next
End Sub

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

Sub SetModLamp(nr, value)
	dim TurnedOn:TurnedOn = 0
	if value > 0 then TurnedOn = 1
	If TurnedOn <> LampState(nr) Then
		LampState(nr) = abs(TurnedOn)
		FadingLevel(nr) = abs(TurnedOn) + 4
	end if 
	ModLevel(nr) = value
End Sub

'//////////////////////////////////////////////////////////////////////
'// VPX STANDARD LIGHTS 
'//////////////////////////////////////////////////////////////////////

Sub NFadeL(nr, object)
Select Case FadingLevel(nr)
Case 4:object.state = 0:FadingLevel(nr) = 0
Case 5:object.state = 1:FadingLevel(nr) = 1
End Select
if (ModLevel(nr) > 0) Then
	object.IntensityScale = ModLevel(nr) / 255
end if 

End Sub

Sub NFadeLm(nr, object, object2) ' used for multiple lights
Select Case FadingLevel(nr)
Case 4:object.state = 0 : object2.state=0
Case 5:object.state = 1: object2.state=1
End Select
if (ModLevel(nr) > 0) Then
	object.IntensityScale = ModLevel(nr) / 128 : object2.IntensityScale=ModLevel(nr) / 128
end if 
End Sub


Sub NFadeLmf(nr, object, object2) ' used for multiple lights
Select Case FadingLevel(nr)
Case 4:object.state = 0 : object2.visible=0
Case 5:object.state = 1: object2.visible=1
End Select
End Sub


Sub NFadeLf(nr, object) ' used for multiple lights
Select Case FadingLevel(nr)
Case 4:object.visible=0
Case 5:object.visible=1
End Select
End Sub



Sub NFadeLmm(nr, object, object2, object3) ' used for multiple lights
Select Case FadingLevel(nr)
Case 4:object.state = 0 : object2.state=0 : object3.state=0
Case 5:object.state = 1: object2.state=1 : object3.state=1
End Select
if (ModLevel(nr) > 0) Then
	object.IntensityScale = ModLevel(nr) / 128 : object2.IntensityScale=ModLevel(nr) / 128 : object3.IntensityScale =ModLevel(nr) / 128
end if 
End Sub

'//////////////////////////////////////////////////////////////////////
'// VPX RAMP & PRIMITIVE LIGHTS 
'//////////////////////////////////////////////////////////////////////

Sub FadeObj(nr, object, a, b, c, d)
Select Case FadingLevel(nr)
Case 4:object.image = b:FadingLevel(nr) = 6                   'fading to off...
Case 5:object.image = a:FadingLevel(nr) = 1                   'ON
Case 6, 7, 8:FadingLevel(nr) = FadingLevel(nr) + 1            'wait
Case 9:object.image = c:FadingLevel(nr) = FadingLevel(nr) + 1 'fading...
Case 10, 11, 12:FadingLevel(nr) = FadingLevel(nr) + 1         'wait
Case 13:object.image = d:FadingLevel(nr) = 0                  'Off
End Select
End Sub

Sub FadeObjm(nr, object, a, b, c, d)
Select Case FadingLevel(nr)
Case 4:object.image = b
Case 5:object.image = a
Case 9:object.image = c
Case 13:object.image = d
End Select
End Sub

Sub NFadeObj(nr, object, a, b)
Select Case FadingLevel(nr)
Case 4:object.image = b:FadingLevel(nr) = 0 'off
Case 5:object.image = a:FadingLevel(nr) = 1 'on
End Select
End Sub


Sub NFadeObjmm(nr, object, object2, a, b, c, d)
Select Case FadingLevel(nr)
Case 4:object.image = b:FadingLevel(nr) = 0: object2.image = c:FadingLevel(nr) = 0 'off 
Case 5:object.image = a:FadingLevel(nr) = 1: object2.image = d:FadingLevel(nr) = 1 'on
End Select
End Sub



Sub NFadeObjm(nr, object, a, b)
Select Case FadingLevel(nr)
Case 4:object.image = b
Case 5:object.image = a
End Select
End Sub

'//////////////////////////////////////////////////////////////////////
'// VPX FLASHER OBJECTS 
'//////////////////////////////////////////////////////////////////////

Sub SetFlash(nr, stat)
FadingLevel(nr) = ABS(stat)
End Sub

Sub Flash(nr, object)
Select Case FadingLevel(nr)
Case 4 'off
FlashLevel(nr) = FlashLevel(nr) - FlashSpeedDown(nr)
If FlashLevel(nr) < FlashMin(nr) Then
FlashLevel(nr) = FlashMin(nr)
FadingLevel(nr) = 0 'completely off
End if
Object.IntensityScale = FlashLevel(nr)
Case 5 ' on
FlashLevel(nr) = FlashLevel(nr) + FlashSpeedUp(nr)
If FlashLevel(nr) > FlashMax(nr) Then
FlashLevel(nr) = FlashMax(nr)
FadingLevel(nr) = 1 'completely on
End if
Object.IntensityScale = FlashLevel(nr)
End Select
End Sub

Sub Flashm(nr, object) 'multiple flashers, it just sets the flashlevel
Object.IntensityScale = FlashLevel(nr)
End Sub

'//////////////////////////////////////////////////////////////////////
'// VPX REELS & TEXT  
'//////////////////////////////////////////////////////////////////////

Sub FadeR(nr, object)
Select Case FadingLevel(nr)
Case 4:object.SetValue 1:FadingLevel(nr) = 6                   'fading to off...
Case 5:object.SetValue 0:FadingLevel(nr) = 1                   'ON
Case 6, 7, 8:FadingLevel(nr) = FadingLevel(nr) + 1             'wait
Case 9:object.SetValue 2:FadingLevel(nr) = FadingLevel(nr) + 1 'fading...
Case 10, 11, 12:FadingLevel(nr) = FadingLevel(nr) + 1          'wait
Case 13:object.SetValue 3:FadingLevel(nr) = 0                  'Off
End Select
End Sub

Sub FadeRm(nr, object)
Select Case FadingLevel(nr)
Case 4:object.SetValue 1
Case 5:object.SetValue 0
Case 9:object.SetValue 2
Case 3:object.SetValue 3
End Select
End Sub

Sub NFadeT(nr, object, message)
Select Case FadingLevel(nr)
Case 4:object.Text = "":FadingLevel(nr) = 0
Case 5:object.Text = message:FadingLevel(nr) = 1
End Select
End Sub

Sub NFadeTm(nr, object, b)
Select Case FadingLevel(nr)
Case 4:object.Text = ""
Case 5:object.Text = message
End Select
End Sub



'//////////////////////////////////////////////////////////////////////
'	FLIPPER SHADOWS
'//////////////////////////////////////////////////////////////////////

Sub GraphicsTimer_Timer()
		batleft.objrotz = LeftFlipper.CurrentAngle + 1
		batleftshadow.objrotz = batleft.objrotz
		batright.objrotz = RightFlipper.CurrentAngle - 1
		batrightshadow.objrotz  = batright.objrotz 		
End Sub

'//////////////////////////////////////////////////////////////////////
'	BALL SHADOWS
'//////////////////////////////////////////////////////////////////////


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
	If UBound(BOT) > UBound (BallShadow) Then Exit Sub
    ' render the shadow for each ball
    For b = 0 to UBound(BOT)
	    if not BOT(b) is Cball then
        If BOT(b).X < 1024/2 Then
            BallShadow(b).X = ((BOT(b).X) - (Ballsize/6) + ((BOT(b).X - (1024/2))/7)) + 5
        Else
            BallShadow(b).X = ((BOT(b).X) + (Ballsize/6) + ((BOT(b).X - (1024/2))/7)) - 5
        End If
        ballShadow(b).Y = BOT(b).Y + 10
        If BOT(b).Z > 20 Then
            BallShadow(b).visible = 1
        Else
            BallShadow(b).visible = 0
        End If
	    end if 
    Next
End Sub



'   _________    ____  ________  __   _       ________    __       ____  ______
'   / ____/   |  / __ \/_  __/ / / /  | |     / /  _/ /   / /      / __ )/ ____/
'  / __/ / /| | / /_/ / / / / /_/ /   | | /| / // // /   / /      / __  / __/   
' / /___/ ___ |/ _, _/ / / / __  /    | |/ |/ // // /___/ /___   / /_/ / /___   
'/_____/_/  |_/_/ |_| /_/ /_/ /_/     |__/|__/___/_____/_____/  /_____/_____/   
'                                                                               
'   ____  __  ______  _____
'  / __ \/ / / / __ \/ ___/
' / / / / / / / /_/ /\__ \ 
'/ /_/ / /_/ / _, _/___/ / 
'\____/\____/_/ |_|/____/  
'
'               \ /                                          \   /
'              --o--           `\\             //'      .____-/.\-____.
'                                \\           //             ~`-'~
'                                 \\. __-__ .//
'                       ___/-_.-.__`/~     ~\'__.-._-\___                    
'.|.       ___________.'__/__ ~-[ \.\'-----'/./ ]-~ __\__`.___________       .|.
'~o~~~~~~~--------______-~~~~~-_/_/ |   .   | \_\_-~~~~~-______--------~~~~~~~o~
'' `               + + +  (X)(X)  ~--\__ __/--~  (X)(X)  + + +               ' `
'                             (X) `/.\' ~ `/.\' (X)  
'                                 "\_/"   "\_/"
'
'							 T   O   P      G   U   N


' Top Gun Table MOD / 4 Players
' By zimbakin 2019, version 1.0b

' MOD Based on JP's Space Shuttle

' ************************************************

' THANKS to:

' - JPSalas for allowing me to mod his table
' - the gurus who took time to answer some rudimentary scripting questions
' - ScottyWic's orbital pinball tutorials & snippets, lindenver's score snippet
' - Nailbuster and team for the awesome PUP system


' PUP/TABLE NOTES:
' ---------------

' TO DO:

'	- Message me if you have script fixes o improvements
'	- Bug Fixes (Game Over PUP, ball lock activation, multiball activation)
'	- Create TILT
' 	- Merge BallinPlay/Contball into 1 (double up)

' ************************************************

Option Explicit
Randomize

' ************************************************************************************************
' Player Options
' ************************************************************************************************

	soundtrackvol = 78 'Set the background audio volume to whatever you'd like out of 100
	videovol = 100 'set the volme you'd like for the videos
	calloutlowermusicvol = 1 'set to 1 if you want music volume lowered during audio callouts

' ************************************************************************************************

	

	' Define Global Variables
	Dim soundtrackvol
	Dim videovol
	Dim calloutlowermusicvol
	Dim BallinPlay



Const BallSize = 50
Const BallMass = 1.1

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

LoadVPM "01550000", "Taito.vbs", 3.26

Dim bsTrough, dtCbank, dtRbank, bsLeftSaucer, bsRightSaucer, x

Const cGameName = "sshuttle"

Const UseSolenoids = 2
Const UseLamps = 0
Const UseGI = 0
Const UseSync = 0 'set it to 1 if the table runs too fast
Const HandleMech = 0

Dim VarHidden
If Table1.ShowDT = true then
    VarHidden = 1
    For each x in aReels
        x.Visible = 1
    Next
else
    VarHidden = 0
    For each x in aReels
        x.Visible = 0
    Next
    lrail.Visible = 0
    rrail.Visible = 0
end if

if B2SOn = true then VarHidden = 1

' Standard Sounds
Const SSolenoidOn = "fx_Solenoidon"
Const SSolenoidOff = "fx_Solenoidoff"
Const SCoin = "fx_coin"



'**************************
'   PinUp Player Config
'   Change HasPuP = True if using PinUp Player Videos
'**************************

	Dim HasPup:HasPuP = True

	Dim PuPlayer
	Dim Scores(26)
	Dim testobject
	
		Const numberfont="digital-7"

		Const pDMD=1
		Const pBackglass=2
		Const pPlayfield=3
		Const pMusic=4
		Const pAudio=7
		Const pCallouts=8

	if HasPuP Then
	on error resume next
	Set PuPlayer = CreateObject("PinUpPlayer.PinDisplay") 
	on error goto 0
	if not IsObject(PuPlayer) then HasPuP = False
	end If

	PuPlayer.Init pBackglass,"TopGun"
	PuPlayer.Init pMusic,"TopGun"
	PuPlayer.Init pCallouts,"TopGun"

	PuPlayer.SetScreenex pBackglass,0,0,0,0,0       'Set TO Always ON    <screen number> , xpos, ypos, width, height, POPUP

	PuPlayer.playlistadd pBackglass,"DefaultBackglass",1,0
	PuPlayer.playlistplayex pBackglass,"DefaultBackglass","DefaultBackglass.mp4",0,1
	PuPlayer.SetBackground pBackglass,1
	
	PuPlayer.SetScreenex pAudio,0,0,0,0,2
	PuPlayer.hide pAudio
	PuPlayer.SetScreenex pMusic,0,0,0,0,2
	PuPlayer.hide pMusic
	PuPlayer.SetScreenex pCallouts,0,0,0,0,2
	PuPlayer.hide pCallouts

	Sub chilloutthemusic 'Lowers volume for PUP videos
		If calloutlowermusicvol = 1 Then
			PuPlayer.SendMSG "{ ""mt"":301, ""SN"": 4, ""FN"":11, ""VL"":40 }"
			ChillTimer.Enabled = True
		End If
	End Sub

Sub ChillTimer_Timer
	ChillTimer.Enabled = False
	turnitbackup
End Sub

Sub chilloutthemusicTG 'Lowers volume for TOPGUN Bonus video
		If calloutlowermusicvol = 1 Then
			PuPlayer.SendMSG "{ ""mt"":301, ""SN"": 4, ""FN"":11, ""VL"":50 }"
			ChillTimerTG.Enabled = True
		End If
	End Sub

Sub ChillTimerTG_Timer
	ChillTimerTG.Enabled = False
	turnitbackup
End Sub
	

	Sub turnitbackup
		If calloutlowermusicvol = 1 Then
			PuPlayer.SendMSG "{ ""mt"":301, ""SN"": 4, ""FN"":11, ""VL"":"&(soundtrackvol)&" }"
		End If
	End Sub

    PuPlayer.LabelInit pBackglass

	'Setup Pages.  Note if you use fonts they must be in FONTS folder of the pupVideos\tablename\FONTS
	'syntax - PuPlayer.LabelNew <screen# or pDMD>,<Labelname>,<fontName>,<size %>,<colour>,<rotation>,<xAlign>,<yAlign>,<xpos>,<ypos>,<PageNum>,<visible>

	'Page 1 (default score display)
	
	PuPlayer.playlistadd pMusic,"MusicGameplay",0,65	 'music rests for 65 seconds before a reset is available
	PuPlayer.playlistadd pMusic,"audioclear",0,0
	PuPlayer.playlistadd pDMD,"DMD",0,0
	PuPlayer.playlistadd pBackglass,"AttractMode",0,0
	PuPlayer.playlistadd pBackglass,"DefaultBackglass",0,0
	PuPlayer.playlistadd pBackglass,"GameOver",0,0
	PuPlayer.playlistadd pBackglass,"MiddleRamp",0,20
	PuPlayer.playlistadd pBackglass,"MultiBall",0,0
	PuPlayer.playlistadd pBackglass,"Multipliers",0,5
	PuPlayer.playlistadd pBackglass,"RightRamp",0,13
	PuPlayer.playlistadd pBackglass,"StartGame",17,30
	PuPlayer.playlistadd pBackglass,"TopGun",0,65
	PuPlayer.playlistadd pBackglass,"BallDrain",0,20
	PuPlayer.playlistadd pBackglass,"BallLock",0,0

	PuPlayer.LabelNew pBackglass,"Play1",numberfont,		4,52224  ,0,0,1,5,85,1,0
	PuPlayer.LabelNew pBackglass,"Play1score",numberfont,	10,59648  ,0,0,1,9,85,1,1
	PuPlayer.LabelNew pBackglass,"Play2",numberfont,		4,52224  ,0,0,1,5,94,1,0
	PuPlayer.LabelNew pBackglass,"Play2score",numberfont,	10,59648  ,0,0,1,9,94,1,1
	PuPlayer.LabelNew pBackglass,"Play3",numberfont,		4,52224  ,0,0,1,74,85,1,0
	PuPlayer.LabelNew pBackglass,"Play3score",numberfont,	10,59648  ,0,0,1,78,85,1,1
	PuPlayer.LabelNew pBackglass,"Play4",numberfont,		4,52224  ,0,0,1,74,94,1,0
	PuPlayer.LabelNew pBackglass,"Play4score",numberfont,	10,59648  ,0,0,1,78,94,1,1
	PuPlayer.LabelNew pBackglass,"TFuel",numberfont,		2,52224  ,0,0,1,48,87,1,0
	PuPlayer.LabelNew pBackglass,"Credits",numberfont,		7,59648  ,0,1,1,52,87,1,1
	PuPlayer.LabelNew pBackglass,"TBall",numberfont,		2,52224  ,0,0,1,54,87,1,0
	PuPlayer.LabelNew pBackglass,"Balls",numberfont,		7,59648  ,0,1,1,58,87,1,1
	PuPlayer.LabelShowPage pBackglass,1,0,""

'Set Background video on DMD
		PuPlayer.playlistplayex pDMD,"DMD","DMD.png",0,1

Sub pUpdateScores
	PuPlayer.LabelSet pBackglass,"Play1","P1>",1,""
	PuPlayer.LabelSet pBackglass,"Play2","P2>",1,""
	PuPlayer.LabelSet pBackglass,"Play3","P3>",1,""
	PuPlayer.LabelSet pBackglass,"Play4","P4>",1,""
	PuPlayer.LabelSet pBackglass,"Play1score",scores(0) & scores(1) & scores(2) & " " & scores(3) & scores(4) & scores(5),1,""
	PuPlayer.LabelSet pBackglass,"Play2score",scores(6) & scores(7) & scores(8) & " " & scores(9) & scores(10) & scores(11),1,""
	PuPlayer.LabelSet pBackglass,"Play3score",scores(12) & scores(13) & scores(14) & " " & scores(15) & scores(16) & scores(17),1,""
	PuPlayer.LabelSet pBackglass,"Play4score",scores(18) & scores(19) & scores(20) & " " & scores(21) & scores(22) & scores(23),1,""
	PuPlayer.LabelSet pBackglass,"TFuel","FUEL",1,""
	PuPlayer.LabelSet pBackglass,"TBall","BALL",1,""
	PuPlayer.LabelSet pBackglass,"Balls",scores(24),1,""
	PuPlayer.LabelSet pBackglass,"Credits",scores(25),1,""

end Sub



'************
' Table init.
'************

Sub table1_Init
    vpmInit me
    With Controller
        .GameName = cGameName
        If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description:Exit Sub
        .SplashInfoLine = "Top Gun" & vbNewLine & "VPX table by zimbakin v.1.0.0"
        .HandleKeyboard = 0
        .ShowTitle = 0
        .ShowDMDOnly = 1
        .ShowFrame = 0
        .HandleMechanics = 0
        .Hidden = VarHidden
        .Games(cGameName).Settings.Value("rol") = 0 '1= rotated display, 0= normal
		.Games(cGameName).Settings.Value("sound") = 0
        '.SetDisplayPosition 0,0, GetPlayerHWnd 'restore dmd window position
        On Error Resume Next
        Controller.SolMask(0) = 0
        vpmTimer.AddTimer 2000, "Controller.SolMask(0)=&Hffffffff'" 'ignore all solenoids - then add the Timer to renable all the solenoids after 2 seconds
        Controller.Run GetPlayerHWnd
        On Error Goto 0
    End With

    ' Nudging
    vpmNudge.TiltSwitch = 30
    vpmNudge.Sensitivity = 3
    vpmNudge.TiltObj = Array(Bumper1, Bumper2, Bumper3, LeftSlingshot, RightSlingshot)

    ' Trough
    Set bsTrough = New cvpmBallStack
    With bsTrough
        .InitSw 0, 1, 11, 21, 0, 0, 0, 0
        .InitKick BallRelease, 90, 4
        .InitExitSnd SoundFX("fx_ballrel", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
        .Balls = 3
    End With

    ' Saucers
    Set bsRightSaucer = New cvpmBallStack
    bsRightSaucer.InitSaucer sw3, 3, 180, 8
    bsRightSaucer.InitExitSnd SoundFX("fx_kicker", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
    bsRightSaucer.KickForceVar = 5

    Set bsLeftSaucer = New cvpmBallStack
    bsLeftSaucer.InitSaucer sw2, 2, 180, 8
    bsLeftSaucer.InitExitSnd SoundFX("fx_kicker", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
    bsLeftSaucer.KickForceVar = 5

    ' Drop targets

    set dtCbank = new cvpmdroptarget
    dtCbank.InitDrop sw43, 43
    dtCbank.initsnd SoundFX("", DOFDropTargets), SoundFX("fx_resetdrop", DOFContactors)
    dtCbank.CreateEvents "dtCbank"
	

    set dtRbank = new cvpmdroptarget
    dtRbank.InitDrop Array(sw4, sw14, sw24), Array(4, 14, 24)
    dtRbank.initsnd SoundFX("", DOFDropTargets), SoundFX("fx_resetdrop", DOFContactors)
    dtRbank.CreateEvents "dtRbank"


    ' Main Timer init
    PinMAMETimer.Interval = PinMAMEInterval
    PinMAMETimer.Enabled = 1

    ' Turn on Gi
    vpmtimer.addtimer 2000, "GiOn '"
    solpost 0
End Sub

Sub table1_Paused:Controller.Pause = 1:End Sub
Sub table1_unPaused:Controller.Pause = 0:End Sub
Sub table1_exit:Controller.Games(cGameName).Settings.Value("sound") = 1:Controller.stop:End Sub


'**********
' Keys
'**********

Sub table1_KeyDown(ByVal Keycode)

if keycode = 46 then ' C Key
     If contball = 1 Then
          contball = 0
     Else
          contball = 1
     End If
End If
if keycode = 48 then 'B Key
     If bcboost = 1 Then
          bcboost = bcboostmulti
     Else
          bcboost = 1
     End If
End If
if keycode = 203 then bcleft = 1 ' Left Arrow
if keycode = 200 then bcup = 1 ' Up Arrow
if keycode = 208 then bcdown = 1 ' Down Arrow
if keycode = 205 then bcright = 1 ' Right Arrow

    If keycode = LeftTiltKey Then Nudge 90, 5:PlaySound SoundFX("fx_nudge", 0), 0, 1, -0.1, 0.25
    If keycode = RightTiltKey Then Nudge 270, 5:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0.1, 0.25
    If keycode = CenterTiltKey Then Nudge 0, 6:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0, 0.25
    If keycode = PlungerKey Then PlaySound "fx_PlungerPull", 0, 1, 0.1, 0.25:Plunger.Pullback
    If keycode = RightFlipperKey Then Controller.Switch(31) = 1
    If vpmKeyDown(keycode)Then Exit Sub
End Sub

Sub table1_KeyUp(ByVal Keycode)

if keycode = 203 then bcleft = 0 ' Left Arrow
if keycode = 200 then bcup = 0 ' Up Arrow
if keycode = 208 then bcdown = 0 ' Down Arrow
if keycode = 205 then bcright = 0 ' Right Arrow

    If keycode = RightFlipperKey Then Controller.Switch(31) = 0
    If vpmKeyUp(keycode)Then Exit Sub
    If keycode = PlungerKey Then PlaySound "fx_plunger", 0, 1, 0.1, 0.25:Plunger.Fire
End Sub


'*********************
' Manual ball control
'*********************


Sub StartControl_Hit()
    Set ControlBall = ActiveBall
    contballinplay = true
	BallinPlay = 1
End Sub

Sub StopControl_Hit()
	contballinplay = false
	BallinPlay = 0
End Sub

Dim bcup, bcdown, bcleft, bcright, contball, contballinplay, ControlBall, bcboost
Dim bcvel, bcyveloffset, bcboostmulti

bcboost = 1 'Do Not Change - default setting
bcvel = 4 'Controls the speed of the ball movement
bcyveloffset = -0.01 'Offsets the force of gravity to keep the ball from drifting vertically on the table, should be negative
bcboostmulti = 3 'Boost multiplier to ball veloctiy (toggled with the B key)

Sub BallControl_Timer()
     If Contball and ContBallInPlay then
          If bcright = 1 Then
               ControlBall.velx = bcvel*bcboost
          ElseIf bcleft = 1 Then
               ControlBall.velx = - bcvel*bcboost
          Else
               ControlBall.velx=0
          End If

         If bcup = 1 Then
              ControlBall.vely = -bcvel*bcboost
         ElseIf bcdown = 1 Then
              ControlBall.vely = bcvel*bcboost
         Else
              ControlBall.vely= bcyveloffset
         End If
     End If
End Sub



'*********
' Switches
'*********

' Slings
Dim LStep, RStep

Sub LeftSlingShot_Slingshot
    PlaySound SoundFX("fx_slingshot", DOFContactors), 0, 1, -0.05, 0.05
    DOF 101, DOFPulse
    LeftSling4.Visible = 1
    Lemk.RotX = 26
    LStep = 0
    vpmTimer.PulseSw 42
    LeftSlingShot.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 1:LeftSLing4.Visible = 0:LeftSLing3.Visible = 1:Lemk.RotX = 14
        Case 2:LeftSLing3.Visible = 0:LeftSLing2.Visible = 1:Lemk.RotX = 2
        Case 3:LeftSLing2.Visible = 0:Lemk.RotX = -10:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

Sub RightSlingShot_Slingshot
    PlaySound SoundFX("fx_slingshot", DOFContactors), 0, 1, 0.05, 0.05
    DOF 102, DOFPulse
    RightSling4.Visible = 1
    Remk.RotX = 26
    RStep = 0
    vpmTimer.PulseSw 41
    RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 1:RightSLing4.Visible = 0:RightSLing3.Visible = 1:Remk.RotX = 14
        Case 2:RightSLing3.Visible = 0:RightSLing2.Visible = 1:Remk.RotX = 2
        Case 3:RightSLing2.Visible = 0:Remk.RotX = -10:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub


'****************************************
' Real Time updates using the GameTimer
'****************************************
'used for all the real time updates

Sub GameTimer_Timer
    RollingUpdate
' add any other real time update subs, like gates or diverters
End Sub


'**************************************************************************************

' BUMPERS

Sub Bumper1_Hit:vpmTimer.PulseSw 35:PlaySound "top_Bumper":PlaySound SoundFX("fx_bumper", DOFContactors), 0, 1, -0.05:li41.duration 1, 50, 0:b01.duration 1, 50, 0:Hud1.duration 1, 50, 0:Hud2.duration 1, 50, 0:Hud3.duration 1, 50, 0:End Sub
Sub Bumper2_Hit:vpmTimer.PulseSw 55:PlaySound "top_Bumper":PlaySound SoundFX("fx_bumper", DOFContactors), 0, 1, 0:li123.duration 1, 50, 0:b02.duration 1, 50, 0:Hud1.duration 1, 50, 0:Hud2.duration 1, 50, 0:Hud3.duration 1, 50, 0:End Sub
Sub Bumper3_Hit:vpmTimer.PulseSw 45:PlaySound "top_Bumper":PlaySound SoundFX("fx_bumper", DOFContactors), 0, 1, -0.025:li40.duration 1, 50, 0:b03.duration 1, 50, 0:Hud1.duration 1, 50, 0:Hud2.duration 1, 50, 0:Hud3.duration 1, 50, 0:End Sub


'**************************************************************************************


' Drain & Saucers
Sub Drain_Hit:Playsound "top_BallDrain":Playsound "fx_drain":bsTrough.AddBall Me:End Sub
Sub sw3_Hit::PlaySound "top_BallLock":PlaySound "fx_kicker_enter", 0, 1, 0, 0.05:bsRightSaucer.AddBall 0:End Sub
Sub sw2_Hit::PlaySound "top_BallLock":PlaySound "fx_kicker_enter", 0, 1, -0.05:bsLeftSaucer.AddBall 0:End Sub


' Rollovers
Sub sw51_Hit:Controller.Switch(51) = 1:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub sw51_UnHit:Controller.Switch(51) = 0:End Sub

Sub sw61_Hit:Controller.Switch(61) = 1:PlaySound "fx_sensor", 0, 1, pan(ActiveBall)
		if scores(24) = 5 And TheGate = 0 And MultiBallOn = False Then
		PlaySound "top_GutterDrain"
		ElseIf TheGate = 0 Then
		PuPlayer.playlistplayex pBackglass,"BallDrain","",100,33
		chilloutthemusic
	End If
End Sub

Sub sw61_UnHit:Controller.Switch(61) = 0:End Sub

Sub sw52_Hit:Controller.Switch(52) = 1:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub sw52_UnHit:Controller.Switch(52) = 0:End Sub

Sub sw62_Hit:Controller.Switch(62) = 1:PlaySound "top_GutterDrain":PlaySound "fx_sensor", 0, 1, pan(ActiveBall)
		if scores(24) = 5 And MultiBallOn = False Then
		PlaySound "top_GutterDrain"
	Else
		PuPlayer.playlistplayex pBackglass,"BallDrain","",100,33
		chilloutthemusic
	End If
End Sub

Sub sw62_UnHit:Controller.Switch(62) = 0:End Sub

Sub sw5_Hit:Controller.Switch(5) = 1:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub sw5_UnHit:Controller.Switch(5) = 0:End Sub

Sub sw15_Hit:Controller.Switch(15) = 1:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub sw15_UnHit:Controller.Switch(15) = 0:End Sub

Sub sw25_Hit:Controller.Switch(25) = 1:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub sw25_UnHit:Controller.Switch(25) = 0:End Sub

Sub sw44_Hit:Controller.Switch(44) = 1:PlaySound "fx_sensor", 0, 1, pan(ActiveBall)
	PuPlayer.playlistplayex pBackglass,"MiddleRamp","",100,16
	chilloutthemusic
End Sub
Sub sw44_UnHit:Controller.Switch(44) = 0:End Sub

Sub sw64_Hit:Controller.Switch(64) = 1:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub sw64_UnHit:Controller.Switch(64) = 0:End Sub


' Right Ramp Exit
Sub sw65_Hit:vpmTimer.PulseSw 65:End Sub

' Droptargets (only sound effect)
Sub sw4_Hit:PlaySound SoundFX("fx_droptarget", DOFDropTargets), 0, 1, pan(ActiveBall):End Sub
Sub sw14_Hit:PlaySound SoundFX("fx_droptarget", DOFDropTargets), 0, 1, pan(ActiveBall):End Sub
Sub sw24_Hit:PlaySound SoundFX("fx_droptarget", DOFDropTargets), 0, 1, pan(ActiveBall):End Sub
Sub sw43_Hit:PlaySound SoundFX("fx_droptarget", DOFDropTargets), 0, 1, pan(ActiveBall) :End Sub


' Trigger (sound only)
Sub Trigger001_Hit:Playsound "top_LeftRamp4":End Sub
Sub Trigger002_Hit:Playsound "top_GutterDrain":End Sub


'****************************************************************************************

' PUP Activations

'****************************************************************************************

Dim pReset(5) 
Dim pStatement(5) 'holds future scripts
Dim FX

for fx=0 to 5
pReset(FX)=0
pStatement(FX)=""
next

DIM pTriggerCounter:pTriggerCounter=pTriggerScript.interval

Sub pTriggerScript_Timer()
for fx=0 to 5 
if pReset(fx)>0 Then 
pReset(fx)=pReset(fx)-pTriggerCounter 
if pReset(fx)<=0 Then
pReset(fx)=0
execute(pStatement(fx))
end if 
End if
next
End Sub

Sub TriggerScript(pTimeMS,pScript)
for fx=0 to 5 
if pReset(fx)=0 Then
pReset(fx)=pTimeMS
pStatement(fx)=pScript
Exit Sub
End If 
next
end Sub


'--------------------------
' Check for PUP Activations
'--------------------------

Dim LockLeft
Dim LockRight
Dim MultiBallOn
Dim MultiBallCount

Dim P1Score
Dim P2Score
Dim P3Score
Dim P4Score


Sub PUPCheck_Timer()
	AttractMode
	Multipliers
	TopGunBonus
	PlayerActivity
End Sub


'--------------------------
' Attract Mode
'--------------------------

Sub AttractMode
		
	if scores(24) = 0 Then
		'PuPlayer.playlistplayex pMusic,"audioclear","clear.mp3",100, 1
		PuPlayer.playlistplayex pBackglass,"AttractMode","AttractMode-1.mp4",100,16
		PuPlayer.SetLoop 2,1
		MultiBallOn = False
		P1Score = 0
		P2Score = 0
		P3Score = 0
		P4Score = 0
	End If
End Sub

'--------------------------
' Game Start & MUSIC
'--------------------------


Sub sw64_Hit
	Controller.Switch(64) = 1

		if scores(24) = 1 Then
		PuPlayer.playlistplayex pMusic,"MusicGameplay","",soundtrackvol,50
		PuPlayer.SetLoop 4,1
		PuPlayer.SetLoop 2,0
		PuPlayer.playlistplayex pBackglass,"StartGame","",100,32
		missileLight1.State = 0
		missileLight2.State = 0
		MultiBallCount = 0
	End If
		
		if scores(24) = 5 Then
		GameContinue
		Else
		PuPlayer.playlistplayex pMusic,"MusicGameplay","",soundtrackvol,50
		PuPlayer.SetLoop 4,1
		PuPlayer.SetLoop 2,0
		MultiBallCount = 0
		' START Game / Launch Ball
		PuPlayer.playlistplayex pBackglass,"StartGame","",100,17
		missileLight1.State = 0
		missileLight2.State = 0
		chilloutthemusic
		MultiBallCount = 0
	
	End If
End Sub


'--------------------------
' Game Continue Players 2+
'--------------------------

Sub GameContinue
	PuPlayer.playresume 4

' START Game / Launch Ball
	PuPlayer.playlistplayex pBackglass,"StartGame","",100,32
	missileLight1.State = 0
	missileLight2.State = 0
	MultiBallCount = 0
	chilloutthemusic	
End Sub


'--------------------------
' Stop Video on launch
'--------------------------

Sub Trigger005_Hit
	PuPlayer.playstop 2 'stops PUP video
	turnitbackup ' resume music

End Sub


'--------------------------
' Multipliers
'--------------------------

Sub Multipliers

		if li103.State = 1 And li20.State = 0 And BallinPlay = 1 And gi1.State = 1 Then
		PuPlayer.playlistplayex pBackglass,"Multipliers","multiplier-2x.mp4",100,5
	End If

		if li20.State = 1 And li21.State = 0 And BallinPlay = 1  And gi1.State = 1 And li103.State = 1 Then
		PuPlayer.playlistplayex pBackglass,"Multipliers","multiplier-3x.mp4",100,6
	End If

		if li21.State = 1 And li22.State = 0  And BallinPlay = 1  And gi1.State = 1 And li20.State = 1 Then
		PuPlayer.playlistplayex pBackglass,"Multipliers","multiplier-4x.mp4",100,7
		chilloutthemusic
	End If

		if li22.State = 1 And li113.State = 0  And BallinPlay = 1  And gi1.State = 1 And li21.State = 1 Then
		PuPlayer.playlistplayex pBackglass,"Multipliers","multiplier-5x.mp4",100,8
		chilloutthemusic
	End If

		if li113.State = 1 And li30.State = 0  And BallinPlay = 1  And gi1.State = 1 And li22.State = 1 Then
		PuPlayer.playlistplayex pBackglass,"Multipliers","multiplier-6x.mp4",100,9
		chilloutthemusic
	End If

		if li30.State = 1 And li103.State = 1 And li20.State = 1 And li21.State = 1 And li22.State = 1 And li113.State = 1 And BallinPlay = 1 And gi1.State = 1 Then
		PuPlayer.playlistplayex pBackglass,"Multipliers","multiplier-7x.mp4",100,10
		chilloutthemusic
	End If

End Sub


'--------------------------
' TOPGUN Bonus
'--------------------------

Sub TopGunBonus
		if BallinPlay = 1 And li112.State = 1 And li111.State = 1 And li110.State = 1 And li121.State = 1 And li120.State = 1 And li119.State = 1 And ChillTimerTG.Enabled = True Then
		PuPlayer.playlistplayex pBackglass,"TopGun","",100,16

		elseIf BallinPlay = 1 And li112.State = 1 And li111.State = 1 And li110.State = 1 And li121.State = 1 And li120.State = 1 And li119.State = 1 And ChillTimerTG.Enabled = False Then
		PuPlayer.playlistplayex pBackglass,"TopGun","",100,16
		chilloutthemusicTG 
	End If
End Sub


'--------------------------
' Ball Locked
'--------------------------

Sub TriggerLL_Hit:PuPlayer.playlistplayex pBackglass,"BallLock","",100,34:LockLeft = 1:End Sub
Sub TriggerLL_UnHit:PuPlayer.playlistplayex pBackglass,"DefaultBackglass","DefaultBackglass.mp4",0,35:LockLeft = 0:End Sub	

Sub TriggerLR_Hit:PuPlayer.playlistplayex pBackglass,"BallLock","",100,34:LockRight = 1:End Sub
Sub TriggerLL_UnHit:PuPlayer.playlistplayex pBackglass,"DefaultBackglass","DefaultBackglass.mp4",0,35:LockRight = 0:End Sub	


'--------------------------
' Multiball
'--------------------------

Sub TriggerMulti_Hit
	If LockLeft = 1 And LockRight = 1 Then
	PuPlayer.playlistplayex pBackglass,"MultiBall","",100,36
	MultiBallOn = True
	FireMissiles
	End if
End Sub

Sub FireMissiles 'Lights up the rockets and chills the music
	missileLight1.State = 2
	missileLight2.State = 2
	chilloutthemusic
End Sub



'--------------------------
' Player Awareness p1,p2...
'--------------------------

Sub PlayerActivity
	P1Score = "0" + scores(0) + scores(2) + scores(3) + scores(4) + scores(4) + scores(5)
	P2Score = "0" + scores(6) + scores(7) + scores(8) + scores(9) + scores(10) + scores(11)
	P3Score = "0" + scores(12) + scores(13) + scores(14) + scores(15) + scores(16) + scores(17)
	P4Score = "0" + scores(18) + scores(19) + scores(20) + scores(21) + scores(22) + scores(23)
End Sub


'--------------------------
' Game Over
'--------------------------

Sub StopControl_Hit

' Count Multiball Drains
		if MultiBallOn = True Then
		MultiBallCount = MultiBallCount + 1
	End If
		if scores(24) = 5 And MultiBallOn = False Then
		CheckPlayers1
	End If
		if scores(24) = 5 And MultiBallOn = True And MultiBallCount <= 2 Then
		Playsound "fx_drain"
	End If
		if scores(24) = 5 And MultiBallOn = True And MultiBallCount >= 3 And Controller.Switch(64) = 0 Then
		CheckPlayers2
	End If

End Sub

Sub CheckPlayers1 'Check if other players are active

		if P2Score <=0 Then
		PuPlayer.playpause 4
		GameIsOver
	End If
		if P3Score <=0 Then
		PuPlayer.playpause 4
		GameIsOver
	End If
		if P4Score <=0 Then
		PuPlayer.playpause 4
		GameIsOver
	End If
		if P4Score >=1 Then
		PuPlayer.playpause 4
		GameIsOver
	End If

End Sub

Sub CheckPlayers2 'Check if other players are active during multiball

		if P2Score <=0 Then
		PuPlayer.playpause 4
		GameIsOver2
	End If	
		if P3Score <=0 Then
		PuPlayer.playpause 4
		GameIsOver2
	End If	
		if P4Score <=0 Then
		PuPlayer.playpause 4
		GameIsOver2
	End If
		if P4Score >=1 Then
		PuPlayer.playpause 4
		GameIsOver2
	End If

End Sub


Sub GameIsOver
		missileLight1.State = 0
		missileLight2.State = 0
		MultiBallOn = False
		MultiBallCount = 0
		PuPlayer.playstop 4
		PuPlayer.SetLoop 2,0
		PuPlayer.playlistplayex pBackglass,"GameOver","",100,31
		TriggerScript 43000,"PriorityClear'"
End Sub

Sub GameIsOver2
		missileLight1.State = 0
		missileLight2.State = 0
		MultiBallOn = False
		MultiBallCount = 0
		PuPlayer.playstop 4
		PuPlayer.SetLoop 2,0
		PuPlayer.playlistplayex pBackglass,"GameOver","",100,37
		TriggerScript 43000,"PriorityClear'"
End Sub

Sub PriorityClear
		'PuPlayer.playlistplayex pBackglass,"DefaultBackglass","VideoClear.mp4",0,50
		PuPlayer.playstop 2
		TriggerScript 9000,"AttractModeOver'"
End Sub
	

Sub AttractModeOver
		PuPlayer.playlistplayex pMusic,"audioclear","clear.mp3",100, 1
		PuPlayer.playlistplayex pBackglass,"AttractMode","AttractMode-1.mp4",100,16
		PuPlayer.SetLoop 2,1
		MultiBallOn = False
		MultiBallCount = 0
		P1Score = 0
		P2Score = 0
		P3Score = 0
		P4Score = 0
End Sub



'****************************************************************************************

' Right Speed Ramp magical ball move... Shazam!

'****************************************************************************************

Sub Kicker1_Hit
  	Kicker1.DestroyBall
	Kicker2.CreateBall
 	Kicker2.Kick -90,20 'Direction and speed
 End Sub


'****************************************************************************************

' Ramp speed boosts

'****************************************************************************************

Sub Trigger003_Hit:ActiveBall.VelX = ActiveBall.VelX*4:PuPlayer.playlistplayex pBackglass,"RightRamp","",100,16:chilloutthemusic:End Sub 'Speed up ball on ramp
Sub Trigger004_Hit:ActiveBall.VelX = ActiveBall.VelX*4:End Sub 'Speed up ball on ramp


'****************************************************************************************

' MOVE F14 Animation 

'****************************************************************************************


Sub sw43_Hit
	PlaySound "top_LockOn6"
	sw43.IsDropped = True
	FighterJet.RotY = 60:FighterJet.RotZ = 100:FighterJet.TransX = 50:FighterJet.TransY = 35:FighterJet.TransZ = 55:FighterJet.ObjRotZ = 21
	gi901.State=0:gi002.State=0:gi003.State=0	'F14 lights off
	gi007.State=1:gi003.State=0	'red cockpit light on
	ThrustL2.State = 1:ThrustL1.State = 2:ThrustL3.State = 1	'Afterburner on
	PlaySound"fx_SolenoidOn"

	End Sub


' ***************************************************************************************



' Spinners
Sub Spinner1_Spin:vpmTimer.PulseSw 71:PlaySound "fx_spinner", 0, 1, 0.05:End Sub

' Targets
Sub sw13_Hit:vpmTimer.PulseSw 13:PlaySound "top_CircTarget3":PlaySound SoundFX("fx_target", DOFDropTargets), 0, 1, pan(ActiveBall):End Sub
Sub sw23_Hit:vpmTimer.PulseSw 23:PlaySound "top_CircTarget3":PlaySound SoundFX("fx_target", DOFDropTargets), 0, 1, pan(ActiveBall):End Sub
Sub sw33_Hit:vpmTimer.PulseSw 33:PlaySound "top_CircTarget3":PlaySound SoundFX("fx_target", DOFDropTargets), 0, 1, pan(ActiveBall):End Sub
Sub sw53_Hit:vpmTimer.PulseSw 53:PlaySound "top_CircTarget3":PlaySound SoundFX("fx_target", DOFDropTargets), 0, 1, pan(ActiveBall):End Sub
Sub sw63_Hit:vpmTimer.PulseSw 63:PlaySound "top_CircTarget3":PlaySound SoundFX("fx_target", DOFDropTargets), 0, 1, pan(ActiveBall):End Sub
Sub sw73_Hit:vpmTimer.PulseSw 73:PlaySound "top_CircTarget3":PlaySound SoundFX("fx_target", DOFDropTargets), 0, 1, pan(ActiveBall):End Sub
Sub sw34_Hit:vpmTimer.PulseSw 34:PlaySound "top_RightTarget3":PlaySound SoundFX("fx_target", DOFDropTargets), 0, 1, pan(ActiveBall):End Sub


'Rubber animations
Dim Rub1, Rub2

Sub sw22_Hit: vpmTimer.PulseSw 22:Rub1 = 1:sw22_Timer:End Sub
Sub sw22_Timer
    Select Case Rub1
        Case 1:r12.Visible = 1:sw22.TimerEnabled = 1
        Case 2:r12.Visible = 0:r13.Visible = 1
        Case 3:r13.Visible = 0:sw22.TimerEnabled = 0
    End Select
    Rub1 = Rub1 + 1
End Sub

Sub sw72_Hit: vpmTimer.PulseSw 72:Rub2 = 1:sw72_Timer:End Sub
Sub sw72_Timer
    Select Case Rub2
        Case 1:r10.Visible = 1:sw72.TimerEnabled = 1
        Case 2:r10.Visible = 0:r11.Visible = 1
        Case 3:r11.Visible = 0:sw72.TimerEnabled = 0
    End Select
    Rub2 = Rub2 + 1
End Sub

' Rubbers
Sub sw12_Hit:vpmTimer.PulseSw 12:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub sw32_Hit:vpmTimer.PulseSw 32:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub sw54_Hit:vpmTimer.PulseSw 54:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub sw72_Hit:vpmTimer.PulseSw 72:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub

'*********
'Solenoids
'*********
SolCallback(1) = "bsTrough.SolOut"
SolCallback(2) = "bsLeftSaucer.SolOut"
SolCallback(3) = "bsRightSaucer.SolOut"
SolCallback(4) = "SetLamp 104,"
SolCallback(5) = "SolPost"
SolCallback(6) = "SolGate"
SolCallback(7) = "ResetdtCbank" 'drop target
SolCallback(8) = "dtRbank.SolDropUp"
SolCallback(17) = "SolGi" '17=relay
SolCallback(18) = "vpmNudge.SolGameOn"


'****************************************************************************************

' RESET F14

'****************************************************************************************

Sub ResetdtCbank(Enabled)
	If Enabled Then
		dtCbank.DropSol_on
		FighterJet.RotY = 85:FighterJet.RotZ = 100:FighterJet.TransX = 0:FighterJet.TransY = 0:FighterJet.TransZ = 0:FighterJet.ObjRotZ = 20
		gi901.State=2:gi002.State=2:gi003.State=1	'F14 lights on
		gi007.State=0:gi003.State=1	'red cockpit light off
		ThrustL2.State = 0:ThrustL1.State = 0:ThrustL3.State = 0	'Afterburner off
		PlaySound"fx_SolenoidOff"
	End If
End Sub

'****************************************************************************************


Sub SolGi(enabled)
    If enabled Then
        GiOff
    Else
        GiOn
    End If
End Sub

Sub SolPost(Enabled)
    If Enabled Then
        PostFlipper.RotateToStart
		Popupwall.IsDropped = 0
		PlaySound"fx_SolenoidOn"
		li42.State=0:li001.State=2
    Else
        PostFlipper.RotateToEnd
		Popupwall.IsDropped = 1
		PlaySound"fx_SolenoidOff"
		li001.State=0
    End If
End Sub

Dim TheGate

Sub SolGate(Enabled)
    If Enabled Then
		PlaySound"fx_SolenoidOn"
        DiverterFlipper.RotateToEnd
		TheGate = 1
    Else
		PlaySound"fx_SolenoidOff"
        DiverterFlipper.RotateToStart
		TheGate = 0
    End If
End Sub




'**************
' Flipper Subs
'**************

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
    If Enabled Then
        PlaySound SoundFX("fx_flipperup", DOFFlippers), 0, 1, -0.1, 0.05
        LeftFlipper.RotateToEnd
    Else
        PlaySound SoundFX("fx_flipperdown", DOFFlippers), 0, 1, -0.1, 0.05
        LeftFlipper.RotateToStart
    End If
End Sub

Sub SolRFlipper(Enabled)
    If Enabled Then
        PlaySound SoundFX("fx_flipperup", DOFFlippers), 0, 1, 0.1, 0.05
        RightFlipper.RotateToEnd
    Else
        PlaySound SoundFX("fx_flipperdown", DOFFlippers), 0, 1, 0.1, 0.05
        RightFlipper.RotateToStart
    End If
End Sub

Sub LeftFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 10, -0.1, 0.25
End Sub

Sub RightFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 10, 0.1, 0.25
End Sub

'*****************
'   Gi Effects
'*****************

Dim OldGiState
OldGiState = -1 'start witht he Gi off

Sub GiON
    For each x in aGiLights
        x.State = 1
    Next
End Sub

Sub GiOFF
    For each x in aGiLights
        x.State = 0
    Next
End Sub

Sub GiEffect(enabled)
    If enabled Then
        For each x in aGiLights
            x.Duration 2, 1000, 1
        Next
    End If
End Sub

Sub GIUpdate
    Dim tmp, obj
    tmp = Getballs
    If UBound(tmp) <> OldGiState Then
        OldGiState = Ubound(tmp)
        If UBound(tmp) = -1 Then
            GiOff
        Else
            GiOn
        End If
    End If
End Sub

'***************************************************
'       JP's VP10 Fading Lamps & Flashers
'       Based on PD's Fading Light System
' SetLamp 0 is Off
' SetLamp 1 is On
' fading for non opacity objects is 4 steps
'***************************************************

Dim LampState(200), FadingLevel(200)
Dim FlashSpeedUp(200), FlashSpeedDown(200), FlashMin(200), FlashMax(200), FlashLevel(200), FlashRepeat(200)

InitLamps()             ' turn off the lights and flashers and reset them to the default parameters
LampTimer.Interval = 10 ' lamp fading speed
LampTimer.Enabled = 1

' Lamp & Flasher Timers

Sub LampTimer_Timer()
    Dim chgLamp, num, chg, ii
    chgLamp = Controller.ChangedLamps
    If Not IsEmpty(chgLamp)Then
        For ii = 0 To UBound(chgLamp)
            LampState(chgLamp(ii, 0)) = chgLamp(ii, 1)       'keep the real state in an array
            FadingLevel(chgLamp(ii, 0)) = chgLamp(ii, 1) + 4 'actual fading step
        Next
    End If
    If VarHidden Then
        UpdateLeds
    End If
    ' some realtime updates
    UpdateLamps
    GIUpdate
    RollingUpdate
    Diverter.RotZ = DiverterFlipper.CurrentAngle
	Popup.Z = PostFlipper.CurrentAngle
End Sub

Sub UpdateLamps()
    NFadeLm 0, li0
    Flash 0, li0a
    NFadeLm 1, li1
    Flash 1, li1a
    NFadeL 10, li10
    NFadeL 100, li100
    NFadeL 101, li101
    NFadeL 102, li102
    NFadeL 103, li103
    NFadeL 109, li109
    NFadeL 11, li11
    NFadeL 110, li110
    NFadeL 111, li111
    NFadeL 112, li112
    NFadeL 113, li113
    NFadeL 119, li119
    NFadeL 12, li12
    NFadeL 120, li120
    NFadeL 121, li121
    NFadeL 122, li122
'   NFadeL 123, li123
    NFadeL 129, li129
    NFadeL 130, li130
    NFadeLm 131, li131a
    NFadeL 131, li131
    NFadeL 132, li132
    NFadeLm 133, li133a
    NFadeL 133, li133
    NFadeL 143, li143
    NFadeL 153, li153
    NFadeL 2, li2
    NFadeL 20, li20
    NFadeL 21, li21
    NFadeL 22, li22
    NFadeL 30, li30
    NFadeL 31, li31
    NFadeL 32, li32
'   NFadeL 40, li40
'   NFadeL 41, li41
    NFadeL 42, li42
    NFadeL 50, li50
    NFadeL 51, li51
    NFadeL 52, li52
    NFadeL 60, li60
    NFadeL 61, li61
    NFadeL 62, li62
    NFadeL 70, li70
    NFadeL 71, li71
    NFadeL 72, li72
    NFadeL 79, li79
    NFadeL 80, li80
    NFadeL 81, li81
    NFadeL 82, li82
    NFadeLm 83, li83
    Flash 83, li83a
    NFadeL 89, li89
    NFadeL 90, li90
    NFadeL 91, li91
    NFadeL 92, li92
    NFadeL 93, li93
    NFadeL 99, li99
    'backdrop lights
    If VarHidden Then
        NFadeL 139, li139
        NFadeL 140, li140
        NFadeL 141, li141
        NFadeL 142, li142
        NFadeL 149, li149
        NFadeL 150, li150
        NFadeL 151, li151
        NFadeL 152, li152
    End If
    'Flasher
    NFadeLm 104, f4
    'NFadeLm 104, f4a
   ' NFadeLm 104, f4b
    NFadeLm 104, f4c
    'NFadeLm 104, f4d
    'NFadeL 104, f4e
End Sub

' div lamp subs

Sub InitLamps()
    Dim x
    For x = 0 to 200
        LampState(x) = 0        ' current light state, independent of the fading level. 0 is off and 1 is on
        FadingLevel(x) = 4      ' used to track the fading state
        FlashSpeedUp(x) = 0.2   ' faster speed when turning on the flasher
        FlashSpeedDown(x) = 0.1 ' slower speed when turning off the flasher
        FlashMax(x) = 1         ' the maximum value when on, usually 1
        FlashMin(x) = 0         ' the minimum value when off, usually 0
        FlashLevel(x) = 0       ' the intensity of the flashers, usually from 0 to 1
        FlashRepeat(x) = 20     ' how many times the flash repeats
    Next
End Sub

Sub AllLampsOff
    Dim x
    For x = 0 to 200
        SetLamp x, 0
    Next
End Sub

Sub SetLamp(nr, value)
    If value <> LampState(nr)Then
        LampState(nr) = abs(value)
        FadingLevel(nr) = abs(value) + 4
    End If
End Sub

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

'Lights, Ramps & Primitives used as 4 step fading lights
'a,b,c,d are the images used from on to off

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

Sub NFadeObjm(nr, object, a, b)
    Select Case FadingLevel(nr)
        Case 4:object.image = b
        Case 5:object.image = a
    End Select
End Sub

' Flasher objects

Sub Flash(nr, object)
    Select Case FadingLevel(nr)
        Case 4 'off
            FlashLevel(nr) = FlashLevel(nr)- FlashSpeedDown(nr)
            If FlashLevel(nr) <FlashMin(nr)Then
                FlashLevel(nr) = FlashMin(nr)
                FadingLevel(nr) = 0 'completely off
            End if
            Object.IntensityScale = FlashLevel(nr)
        Case 5 ' on
            FlashLevel(nr) = FlashLevel(nr) + FlashSpeedUp(nr)
            If FlashLevel(nr)> FlashMax(nr)Then
                FlashLevel(nr) = FlashMax(nr)
                FadingLevel(nr) = 1 'completely on
            End if
            Object.IntensityScale = FlashLevel(nr)
    End Select
End Sub

Sub Flashm(nr, object) 'multiple flashers, it doesn't change anything, it just follows the main flasher
    Select Case FadingLevel(nr)
        Case 4, 5
            Object.IntensityScale = FlashLevel(nr)
    End Select
End Sub

Sub FlashBlink(nr, object)
    Select Case FadingLevel(nr)
        Case 4 'off
            FlashLevel(nr) = FlashLevel(nr)- FlashSpeedDown(nr)
            If FlashLevel(nr) <FlashMin(nr)Then
                FlashLevel(nr) = FlashMin(nr)
                FadingLevel(nr) = 0 'completely off
            End if
            Object.IntensityScale = FlashLevel(nr)
            If FadingLevel(nr) = 0 AND FlashRepeat(nr)Then 'repeat the flash
                FlashRepeat(nr) = FlashRepeat(nr)-1
                If FlashRepeat(nr)Then FadingLevel(nr) = 5
            End If
        Case 5 ' on
            FlashLevel(nr) = FlashLevel(nr) + FlashSpeedUp(nr)
            If FlashLevel(nr)> FlashMax(nr)Then
                FlashLevel(nr) = FlashMax(nr)
                FadingLevel(nr) = 1 'completely on
            End if
            Object.IntensityScale = FlashLevel(nr)
            If FadingLevel(nr) = 1 AND FlashRepeat(nr)Then FadingLevel(nr) = 4
    End Select
End Sub

' Desktop Objects: Reels & texts (you may also use lights on the desktop)

' Reels

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

'Texts

Sub NFadeT(nr, object, message)
    Select Case FadingLevel(nr)
        Case 4:object.Text = "":FadingLevel(nr) = 0
        Case 5:object.Text = message:FadingLevel(nr) = 1
    End Select
End Sub

Sub NFadeTm(nr, object, message)
    Select Case FadingLevel(nr)
        Case 4:object.Text = ""
        Case 5:object.Text = message
    End Select
End Sub

'************************************
'          LEDs Display
'     Based on Scapino's LEDs
'************************************

Dim Digits(32)
Dim Patterns(11)
Dim Patterns2(11)

Patterns(0) = 0     'empty
Patterns(1) = 63    '0
Patterns(2) = 6     '1
Patterns(3) = 91    '2
Patterns(4) = 79    '3
Patterns(5) = 102   '4
Patterns(6) = 109   '5
Patterns(7) = 125   '6
Patterns(8) = 7     '7
Patterns(9) = 127   '8
Patterns(10) = 111  '9

Patterns2(0) = 128  'empty
Patterns2(1) = 191  '0
Patterns2(2) = 134  '1
Patterns2(3) = 219  '2
Patterns2(4) = 207  '3
Patterns2(5) = 230  '4
Patterns2(6) = 237  '5
Patterns2(7) = 253  '6
Patterns2(8) = 135  '7
Patterns2(9) = 255  '8
Patterns2(10) = 239 '9

'Assign 6-digit output to reels
Set Digits(0) = a0
Set Digits(1) = a1
Set Digits(2) = a2
Set Digits(3) = a3
Set Digits(4) = a4
Set Digits(5) = a5

Set Digits(6) = b0
Set Digits(7) = b1
Set Digits(8) = b2
Set Digits(9) = b3
Set Digits(10) = b4
Set Digits(11) = b5

Set Digits(12) = c0
Set Digits(13) = c1
Set Digits(14) = c2
Set Digits(15) = c3
Set Digits(16) = c4
Set Digits(17) = c5

Set Digits(18) = d0
Set Digits(19) = d1
Set Digits(20) = d2
Set Digits(21) = d3
Set Digits(22) = d4
Set Digits(23) = d5

Set Digits(24) = e0
Set Digits(25) = e1

Sub UpdateLeds
    On Error Resume Next
    Dim ChgLED, ii, jj, chg, stat, statdupe,num
    ChgLED = Controller.ChangedLEDs(&HFF, &HFFFF)
    If Not IsEmpty(ChgLED)Then
        For ii = 0 To UBound(ChgLED)
            num=chgLED(ii,0):chg = chgLED(ii, 1):stat = chgLED(ii, 2):statdupe = chgLED(ii, 2)
            For jj = 0 to 10
                If stat = Patterns(jj)OR stat = Patterns2(jj)then Digits(chgLED(ii, 0)).SetValue jj
            Next
			Select Case (statdupe and 255)
				Case 0   : scores(num) = " "
				Case 63  : scores(num) = "0"
				Case 6   : scores(num) = "1"
				Case 91  : scores(num) = "2"
				Case 79  : scores(num) = "3"
				Case 102 : scores(num) = "4"
				Case 109 : scores(num) = "5"
				Case 125 : scores(num) = "6"
				Case 7   : scores(num) = "7"
				Case 127 : scores(num) = "8"
				Case 111 : scores(num) = "9"
				Case 128 : scores(ii) = " ":debug.print " "
				Case 191 : scores(ii) = "0":debug.print "0"
				Case 134 : scores(ii) = "1":debug.print "1"
				Case 219 : scores(ii) = "2":debug.print "2"
				Case 207 : scores(ii) = "3":debug.print "3"
				Case 237 : scores(ii) = "4":debug.print "4"
				Case 230 : scores(ii) = "5":debug.print "5"
				Case 253 : scores(ii) = "6":debug.print "6"
				Case 135 : scores(ii) = "7":debug.print "7"
				Case 255 : scores(ii) = "8":debug.print "8"
				Case 239 : scores(ii) = "9":debug.print "9"
			End Select
			If num=5 Then
				TestObject=statdupe
			End If
        Next
    End IF
	pUpdateScores

'check for PUP activations

'AttractMode
'Multipliers


End Sub


'******************************
' Diverse Collection Hit Sounds
'******************************

Sub aMetals_Hit(idx):PlaySound "fx_MetalHit", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
Sub aRubber_Bands_Hit(idx):PlaySound "fx_rubber_band", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
Sub aRubber_Posts_Hit(idx):PlaySound "fx_rubber_post", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
Sub aRubber_Pins_Hit(idx):PlaySound "fx_rubber_pin", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
Sub aPlastics_Hit(idx):PlaySound "fx_PlasticHit", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
Sub aGates_Hit(idx):PlaySound "fx_Gate", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
Sub aWoods_Hit(idx):PlaySound "fx_Woodhit", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub

' *********************************************************************
'                      Supporting Ball & Sound Functions
' *********************************************************************

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
    Vol = Csng(BallVel(ball) ^2 / 2000)
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "table1" is the name of the table
    Dim tmp
    tmp = ball.x * 2 / table1.width-1
    If tmp> 0 Then
        Pan = Csng(tmp ^10)
    Else
        Pan = Csng(-((- tmp) ^10))
    End If
End Function

Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
    Pitch = BallVel(ball) * 20
End Function

Function BallVel(ball) 'Calculates the ball speed
    BallVel = INT(SQR((ball.VelX ^2) + (ball.VelY ^2)))
End Function

Function AudioFade(ball) 'only on VPX 10.4 and newer
    Dim tmp
    tmp = ball.y * 2 / Table1.height-1
    If tmp> 0 Then
        AudioFade = Csng(tmp ^10)
    Else
        AudioFade = Csng(-((- tmp) ^10))
    End If
End Function

'*****************************************
'      JP's VP10 Rolling Sounds
'*****************************************

Const tnob = 20 ' total number of balls
Const lob = 0   'number of locked balls
ReDim rolling(tnob)
InitRolling

Sub InitRolling
    Dim i
    For i = 0 to tnob
        rolling(i) = False
    Next
End Sub

Sub RollingUpdate()
    Dim BOT, b, ballpitch
    BOT = GetBalls

    ' stop the sound of deleted balls
    For b = UBound(BOT) + 1 to tnob
        rolling(b) = False
        StopSound("fx_ballrolling" & b)
    Next

    ' exit the sub if no balls on the table
    If UBound(BOT) = lob - 1 Then Exit Sub 'there no extra balls on this table

    ' play the rolling sound for each ball
    For b = lob to UBound(BOT)
        If BallVel(BOT(b))> 1 Then
            If BOT(b).z <30 Then
                ballpitch = Pitch(BOT(b))
            Else
                ballpitch = Pitch(BOT(b)) + 15000 'increase the pitch on a ramp or elevated surface
            End If
            rolling(b) = True
            PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b)), Pan(BOT(b)), 0, ballpitch, 1, 0
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
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 2000, Pan(ball1), 0, Pitch(ball1), 0, 0
End Sub


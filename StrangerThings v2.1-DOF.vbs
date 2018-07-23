 '   -::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::-      
 '    `-/+++//:-.---------.--------.`    ----.  `-----. .--``-::::---  .--------`.-sssssssssss-.       
 '   .oss+-.-:oo/o:/sss:/o.osso-:osso   `ssss/   +ssss+`.s::oss:--:+o` .sss+--:+/ +ssss/:+ssss/      
 '   +sss-    `--` -sss` .-+ss+ `+ss/   /o+sss`  ++:oss+.s/sss-     -` `sss/../ : /sss+  `osss/      
 '   +sssso/:.`    -sss`   +sso/oso.   .s-.sss/  ++ .+ssss/sso   `:///-.oss+:/o`  /sss+ `/sso:`      
 '   `+sssssssso:` -sss`   +ss+:sss+` `ooo+ssss. ++  `:sss/sso    -sso``oss:  .` `/ssso+ssso.        
 '     `-:+osssss/ -sss`   +ss/ -sss+ :s` `:sss+ ++    .os:sss:   -sso``oss:    .:+sss+-ossso.       
 '   --     `/ssso :sss.   +ss+  /sss/s/    +sss-++     -s:.+ss/-.+sso`.sss+..-/o`+sss/ .ossss.      
 '   /o-    `/sss-.-:::-` .:::-. -:::::-`  `-:::---.    -:-` `.-::::::.-::::::::- +sss/  .ossso`     
 '   +sso+++oso/.   ````````.`                                ``..```      ```   `ossso.  /ssss+`    
 '   ````....`      oo/ssso/o/:ooo/``/ooo:`:ooo:`/oooo+`.++.:oss+//os-  `:oo++oo-....... `.......    
 '   :////////////-.-  sss: `:.sss:``:sss. -sss. -s+sss/`++/sso.    -: `oss/  `::://////////////:    
 '                     sss:   .sss+//+sss. -sss. -o.-ossoo++ss:   .----.osss+/-.`                    
 '                     sss:   .sss-  -sss. -sss. -o` `+sss++ss-   `oss: .+ossssso.                   
 '                     sss:   .sss-  -sss. -sss. -o`   -os++ss+    oss:-` ``-/sss+                   
 '                     sss/   .sss-  -sss. -sss. :o`    .s+`+ss/```oss::+-` `-sss-                   
 '                   `-+++/.  -+++:``:+++-`:+++:`/+-    .++.`./////+++/-+++++++:.                    
 '*************************************************************************************************                                                                        ````                              
 ' Original Pinball Table Created by ScottyWic
 ' Dedicated to my wife, without her consistent disapproval, there's no way I would have finished.
 ' Thanks to JPSalas for all his instruction on the scripting
 ' Thanks to JrSprad for so many hours of helping me collect scenes for the DMD
 ' Thanks to STAT for the DirectB2s
 ' Thanks to FreeLunch, MnHotRod and lodger for beta testing.
 ' Thanks to Netflix and all who worked on making this unbelievable show.
 ' I hope you enjoy, and please kill the demogorgon won't you?
 ' All Art and resources have been sourced online and we hold no rights to any media used.
 '*************************************************************************************************                                                          
 ' TO-DO's
 ' ************************************************************************************************ 
 ' Release Notes
 ' v2.1 - extra ball, 1 per ball and harder, lane drain removed you need to full drain to get your ball save.. sorry.
 ' v2.0 - I've missed some here but I'll try to remember what I can, there was alot.
 ' Orbital rom config built and added, b2s light show, BG out audio for DOF peeps
 ' ball save timer doesn't start until actaully launched, first place is grand champion, out lanes activate ball save and don't have to wait for drain,
 ' auto plunge on multi & ball save given a defined strength, nancy & steve clips randomized on hit, extra ball requirements dropped,
 ' high scores reworked, game over scene reworked & skippable.intro scenes are all skippable now in a loop
 ' v1.91 - Included Arngrim's Dialed in DOF, and fixed the DOF close - also from Arngrim
 ' v1.9
 ' slingshot band color, extra ball rework, upside down left flipper stall fix, right lower rail fix, lowered height of right ramp,
 ' DOF code by Dieter added, many hit flashers added, movement added to hawkins truck, dice and waffles, 
 ' backwall flashers added, center target light effect, lane sound change to bloop, b2s stop controller on table exit, 
 ' ramp flashers added,
 ' v1.8 - Optimized/Individualized all rewards for 4 player mode, plunger sounds
 ' v1.7 - increased default high scores, fixed Will Super Jackpot, Fixed Extra balls, moved up drain walls to see wall exiting, sad =(
 ' v1.6 - magnets issue with locks, fixed default material items, fresh waffle bumpers - thanks BorgDog
 ' 
 ' v1.5 - fixed Multiball restart super jackpot bug & new game lock available bug
 ' v1.4 - added new toys and fixed some dmd text
 ' v1.3 - fixed will escape lights reset after will multiball started, created delay for cinematics,
 '  added in dmd flushers to make sure important scenes aren't blocked
 '  defaulted game to 3 balls. Can adjust in the script if desired
 '  made right ramp extra ball 2 shots less - 7
 '  added in blocker for upside down w/ other multiballs going 
 ' v1.2 - chilled our the orbit sounds, added Basic DOF calls (no flashforms yet), integrated b2s calls, darkened lower playfield & deactivated flippers until ready         
 ' v1.1 - Gate to upsidedown fixed. Was getting locked after multiballs.  
 ' v1.0 - Initial Release
                                     
Option Explicit
Randomize

'---------- UltraDMD Unique Table Color preference -------------
Dim DMDColor, DMDColorSelect, UseFullColor
Dim DMDPosition, DMDPosX, DMDPosY, DMDSize, DMDWidth, DMDHeight 


UseFullColor = "True" '                           "True" / "False"
DMDColorSelect = "Red"            ' Rightclick on UDMD window to get full list of colours

DMDPosition = True                               ' Use Manual DMD Position, True / False
DMDPosX = 100                                   ' Position in Decimal
DMDPosY = 40                                     ' Position in Decimal

DMDSize = True                                     ' Use Manual DMD Size, True / False
DMDWidth = 512                                    ' Width in Decimal
DMDHeight = 128                                   ' Height in Decimal 

'Note open Ultradmd and right click on window to get the various sizes in decimal 

GetDMDColor
Sub GetDMDColor
Dim WshShell,filecheck,directory
Set WshShell = CreateObject("WScript.Shell")
If DMDSize then
WshShell.RegWrite "HKCU\Software\UltraDMD\w",DMDWidth,"REG_DWORD"
WshShell.RegWrite "HKCU\Software\UltraDMD\h",DMDHeight,"REG_DWORD"
End if
If DMDPosition then
WshShell.RegWrite "HKCU\Software\UltraDMD\x",DMDPosX,"REG_DWORD"
WshShell.RegWrite "HKCU\Software\UltraDMD\y",DMDPosY,"REG_DWORD"
End if
WshShell.RegWrite "HKCU\Software\UltraDMD\fullcolor",UseFullColor,"REG_SZ"
WshShell.RegWrite "HKCU\Software\UltraDMD\color",DMDColorSelect,"REG_SZ"
End Sub
'---------------------------------------------------


Const BallSize = 50 ' 50 is the normal size
Const cGameName = "stranger_things"

' Load the core.vbs for supporting Subs and functions
LoadCoreFiles

Sub LoadCoreFiles
    On Error Resume Next
    ExecuteGlobal GetTextFile("core.vbs")
    If Err Then MsgBox "Can't open core.vbs"
    ExecuteGlobal GetTextFile("controller.vbs")
    If Err Then MsgBox "Can't open controller.vbs"
    On Error Goto 0
End Sub

 ' //////////////////////
 ' B2S Light Show
 ' cause i mean everyone loves a good light show
 ' 1 = left bike Light
 ' 2 = middle bike light 
 ' 3 = right bike light 
 ' 4 = hazmat suit guide
 ' 5 = top left mom
 ' 6 = top right officer
 ' 7 = logo
 ' /////////////////////

Dim b2sstep
b2sstep = 0
b2sflash.enabled = 0
Dim b2satm

Sub startB2S(aB2S)
	b2sflash.enabled = 1
	b2satm = ab2s
End Sub

Sub b2sflash_timer
    If B2SOn Then
	b2sstep = b2sstep + 1
	Select Case b2sstep
		Case 0
		Controller.B2SSetData b2satm, 0
		Case 1
		Controller.B2SSetData b2satm, 1
		Case 2
		Controller.B2SSetData b2satm, 0
		Case 3
		Controller.B2SSetData b2satm, 1
		Case 4
		Controller.B2SSetData b2satm, 0
		Case 5
		Controller.B2SSetData b2satm, 1
		Case 6
		Controller.B2SSetData b2satm, 0
		Case 7
		Controller.B2SSetData b2satm, 1
		Case 8
		Controller.B2SSetData b2satm, 0
		b2sstep = 0
		b2sflash.enabled = 0
	End Select
    End If
End Sub

' /////////////////
' END b2s 
' well that was quick, tacos anyone?
' /////////////////


' Define any Constants
Const TableName = "StrangerThings"
Const myVersion = "1.0.0"
Const MaxPlayers = 4
' Const BallSaverTime = 15 'in seconds
Const MaxMultiplier = 10 'limit to 10x in this game
Const MaxMultiballs = 5  ' max number of balls during multiballs

' Define Global Variables
Dim PlayersPlayingGame
Dim CurrentPlayer
Dim Credits
Dim BonusPoints(4)
Dim BonusHeldPoints(4)
Dim BonusMultiplier(4)
Dim bBonusHeld
Dim BallsRemaining(4)
Dim ExtraBallsAwards(4)
Dim Score(4)
Dim HighScore(4)
Dim HighScoreName(4)
Dim Jackpot
Dim SuperJackpot
Dim Tilt
Dim TiltSensitivity
Dim Tilted
Dim TotalGamesPlayed
Dim mBalls2Eject
Dim SkillshotValue(4)
Dim bAutoPlunger
Dim bInstantInfo
Dim bromconfig
Dim bAttractMode

' Define Game Control Variables
Dim LastSwitchHit
Dim BallsOnPlayfield
Dim BallsInLock(4)
Dim BallsInHole

' Define Game Flags
Dim bFreePlay
Dim bGameInPlay
Dim bOnTheFirstBall
Dim bBallInPlungerLane
Dim bBallSaverActive
Dim bBallSaverReady
Dim bMultiBallMode
Dim bMusicOn
Dim bSkillshotReady
Dim bExtraBallWonThisBall
Dim bJustStarted

Dim plungerIM 'used mostly as an autofire plunger
Dim MagnetR
Dim MagnetU
Dim MagnetB

' *********************************************************************
'                Visual Pinball Defined Script Events
' *********************************************************************

Sub Table1_Init()
	Loadbpg
	If speedcurrent = "Slow" Then
	table1.SlopeMin = 5
	table1.SlopeMax = 5
	End If
	If speedcurrent = "Flow" Then
	table1.SlopeMin = 8
	table1.SlopeMax = 8
	End If

    LoadEM
    Dim i
    Randomize
    'Impulse Plunger as autoplunger
    Const IMPowerSetting = 45 ' Plunger Power
    Const IMTime = 1.1        ' Time in seconds for Full Plunge
    Set plungerIM = New cvpmImpulseP
    With plungerIM
        .InitImpulseP swplunger, IMPowerSetting, IMTime
        .Random 1.5
        .InitExitSnd SoundFX("fx_kicker", DOFContactors), SoundFX("fx_solenoid", DOFContactors)
        .CreateEvents "plungerIM"
    End With


 Set MagnetR = New cvpmMagnet
    With MagnetR
        .InitMagnet Magnet11, 60
        .GrabCenter = True
        .MagnetOn = 0
        .CreateEvents "MagnetR"
    End With


 Set MagnetU = New cvpmMagnet
    With MagnetU
        .InitMagnet MagnetEscape, 70
        .GrabCenter = True
        .MagnetOn = 0
        .CreateEvents "MagnetU"
    End With

 Set MagnetB = New cvpmMagnet
    With MagnetB
        .InitMagnet MagnetBarb, 60
        .GrabCenter = True
        .MagnetOn = 0
        .CreateEvents "MagnetB"
    End With

    'load saved values, highscore, names, jackpot
    Loadhs

    'Init main variables

    ' initalise the DMD display
    DMD_Init

    ' freeplay or coins
   ' bFreePlay = gamemodecurrent 'we dont want coins

    ' initialse any other flags
	Loadbpg
    bAttractMode = False
    bOnTheFirstBall = False
    bBallInPlungerLane = False
    bBallSaverActive = False
    bBallSaverReady = False
    bMultiBallMode = False
    bGameInPlay = False
    bAutoPlunger = False
    bMusicOn = True
    BallsOnPlayfield = 0
For i = 0 to 4
    BallsInLock(i) = 0
Next
    BallsInHole = 0
    LastSwitchHit = ""
    Tilt = 0
	Saves = 0
	Drains = 0
	RRHits(CurrentPlayer) = 0
	LRHits(CurrentPlayer) = 0
    TiltSensitivity = 6
    Tilted = False
    bBonusHeld = False
    bJustStarted = True
    bInstantInfo = False
	bromconfig = False
	bumps(CurrentPlayer) = 0
    RomConfigTimer.Enabled = False
    ' set any lights for the attract mode
    GiOff
    StartAttractMode

    ' Remove the cabinet rails if in FS mode
    If Table1.ShowDT = False then
        lrail.Visible = False
        rrail.Visible = False
    End If
End Sub
	FlasherEL.opacity = 0
	Flasher1.opacity = 100
'******
' Keys
'******

Sub Table1_KeyDown(ByVal Keycode)
	If Keycode = AddCreditKey Then
        Credits = Credits + 1
        DOF 140, DOFOn
        If(Tilted = False) Then
            DMDFlush
            DMD "black.png", "CREDITS " &credits, "PRESS START",  2000
            PlaySound "fx_coin"
            If NOT bGameInPlay Then ShowTableInfo
        End If
    End If

    If keycode = PlungerKey Then
		PlaySound "fx_plungerpull"
        Plunger.Pullback
    End If

    If hsbModeActive Then
        EnterHighScoreKey(keycode)
        Exit Sub
    End If

    If inconfig = True Then
        ConfigKey(keycode)
        Exit Sub
    End If


    ' Table specific

    ' Normal flipper action




    If bGameInPlay Then
		If NOT Tilted Then
	    If keycode = LeftTiltKey Then Nudge 90, 6:PlaySound SoundFX("fx_nudge",0), 0, 1, -0.1, 0.25:CheckTilt
        If keycode = RightTiltKey Then Nudge 270, 6:PlaySound SoundFX("fx_nudge",0), 0, 1, 0.1, 0.25:CheckTilt
        If keycode = CenterTiltKey Then Nudge 0, 7:PlaySound SoundFX("fx_nudge",0), 0, 1, 1, 0.25:CheckTilt

        If keycode = LeftFlipperKey Then SolLFlipper 1:SolULFlipper 1:InstantInfoTimer.Enabled = True
        If keycode = RightFlipperKey Then SolRFlipper 1:SolURFlipper 1:InstantInfoTimer.Enabled = True

        If keycode = StartGameKey Then
            If((PlayersPlayingGame <MaxPlayers) AND(bOnTheFirstBall = True) ) Then

                If(gamemodecurrent = "True") Then
                    PlayersPlayingGame = PlayersPlayingGame + 1
                    TotalGamesPlayed = TotalGamesPlayed + 1
                    DMDFlush
                    DMD "black.png", " ", PlayersPlayingGame & " PLAYERS",  500
                    PlaySound "so_fanfare1"
                Else
                    If(Credits> 0) then
                        PlayersPlayingGame = PlayersPlayingGame + 1
                        TotalGamesPlayed = TotalGamesPlayed + 1
                        Credits = Credits - 1
						If Credits < 1 Then DOF 140, DOFOff
                        DMDFlush
                        DMD "black.png", " ", PlayersPlayingGame & " PLAYERS",  500
                        PlaySound "so_fanfare1"
                    Else
                        ' Not Enough Credits to start a game.
                        DMDFlush
                        DMD "black.png", "CREDITS " &credits, "INSERT COIN",  500
                        PlaySound "so_nocredits"
                    End If
                End If
            End If
        End If
		End If
        Else
		If NOT Tilted Then
 ' If (GameInPlay)
				'If keycode = RightFlipperKey Then DMDFlush
			If keycode = LeftFlipperKey Then SolLFlipper 0:SolULFlipper 0:RomConfigTimer.Enabled = True:dmdintroloop
			If keycode = RightFlipperKey Then SolRFlipper 0:SolURFlipper 0:RomConfigTimer.Enabled = True:dmdintroloop
			If inconfig = False Then
				If keycode = StartGameKey Then
					If(gamemodecurrent = "True") Then
						If(BallsOnPlayfield = 0) Then
							ResetForNewGame()
						End If
					Else
						If(Credits> 0) Then
							If(BallsOnPlayfield = 0) Then
								Credits = Credits - 1
								If Credits < 1 Then DOF 140, DOFOff
								ResetForNewGame()
							End If
						Else
							' Not Enough Credits to start a game.
							DOF 140, DOFOff
							DMDFlush
							DMD "black.png", "CREDITS " &credits, "INSERT COIN",  500
							ShowTableInfo
						End If
					End If
				End If
			Else
			ConfigKey(keycode)
			End If
		End If
		End If ' If (GameInPlay)

'****************
' Testing Keys  
'****************
if keycode = "3" then
AddScore 2000000
CheckHighscore()
End If
End Sub

Sub Table1_KeyUp(ByVal keycode)

    If keycode = PlungerKey Then
		PlaySound "fx_plunger"
        Plunger.Fire
    End If

    If hsbModeActive Then
        Exit Sub
    End If

    ' Table specific

    If bGameInPLay Then
        If keycode = LeftFlipperKey Then
            SolLFlipper 0
			SolULFlipper 0
            InstantInfoTimer.Enabled = False
            If bInstantInfo Then
                DMDScoreNow
                bInstantInfo = False
            End If
        End If
        If keycode = RightFlipperKey Then
            SolRFlipper 0
			SolURFlipper 0 
            InstantInfoTimer.Enabled = False
            If bInstantInfo Then
                DMDScoreNow
                bInstantInfo = False
            End If
        End If
	Else
		If keycode = LeftFlipperKey Then SolLFlipper 0:SolULFlipper 0:RomConfigTimer.Enabled = False
        If keycode = RightFlipperKey Then SolRFlipper 0:SolURFlipper 0:RomConfigTimer.Enabled = False
    End If
End Sub

Sub InstantInfoTimer_Timer
    InstantInfoTimer.Enabled = False
    bInstantInfo = True
    DMDFlush
    UltraDMDTimer.Enabled = 1
End Sub

Sub RomConfigTimer_Timer
	If inconfig = True Then
	StartAttractMode()
	Else
    RomConfigTimer.Enabled = False
	romconfig
	End If
End Sub

Sub InstantInfo
	If finalflips = False Then
    DMD "start-impatient-2.wmv", "", "",  3000
    DMD "black.png", "", "INSTANT INFO",  500
    DMD "black.png", "Left Ramp Hits", LRHits(CurrentPlayer),  1000
    DMD "black.png", "Right Ramp Hits", RRHits(CurrentPlayer),  1000
    DMD "black.png", "Barb Locks", BallsInLock(CurrentPLayer),  1000
    DMD "black.png", "Run Locks", BallsInRunLock(CurrentPlayer),  1000
    DMD "black.png", "Barb Jackpots", barbHits(CurrentPlayer),  1000
    DMD "black.png", "Run Jackpots", RunHits(CurrentPlayer),  1000
    DMD "black.png", "Will Jackpots", WillHits(CurrentPlayer),  1000
    DMD "black.png", "Steve Hits", SteveHit(CurrentPlayer),  1000
    DMD "black.png", "Nancy Hits", NancyHit(CurrentPlayer),  1000
    DMD "black.png", "Eggos Collected", bumps(CurrentPlayer),  1000
    DMD "black.png", "Orbits Hit", orbithit(CurrentPlayer),  1000
	End If
End Sub

'***********************
' Orbital Rom Config
'***********************
Dim configmode
configmode = 0
Dim inconfig
inconfig = false
Sub romconfig
	Loadbpg
	inconfig = true
	DMDFlush
	StopAttractMode()
    DMD "config.wmv", "", "",  4000
    vpmtimer.addtimer 4000, "configoptions '"
End Sub

Sub ConfigKey(keycode)
    If keycode = LeftFlipperKey Then
        Playsound "fx_Previous"
		If mainconfig = True Then
        configoptions
		End If
		If subconfig = true Then
			If cfg1 = true Then
			cfgballs
			End If

			If cfg2 = true Then
			cfgbst
			End If

			If cfg3 = true Then
			cfgfree
			End If

			If cfg4 = true Then
			cfghs
			End If

			If cfg5 = true Then
			cfgprofanity
			End If

			If cfg6 = true Then
			cfgspeed
			End If
		End If
    End If

    If keycode = RightFlipperKey Then
        Playsound "fx_Next"
		If mainconfig = True Then
        configoptions
		End If
		If subconfig = true Then
			If cfg1 = true Then
			cfgballs
			End If

			If cfg2 = true Then
			cfgbst
			End If

			If cfg3 = true Then
			cfgfree
			End If

			If cfg4 = true Then
			cfghs
			End If

			If cfg5 = true Then
			cfgprofanity
			End If

			If cfg6 = true Then
			cfgspeed
			End If
		End If
    End If

    If keycode = StartGameKey Then
	DMDFLush
	DMD "black.png", "Saving", "Config",  999999
	vpmtimer.addtimer 2000, "configexit '"
    end if

    If keycode = PlungerKey OR Keycode = AddCreditKey Then
		If subconfig = true Then
			if cfg1 = true Then
			bpgvalue
			DMDFLush
			DMD "cfg.png", "Saving", "Balls Per Game",  999999
			vpmtimer.addtimer 2000, "configoptions '"
			subconfig = False
			bpg5 = False
			bpg3 = False 
			end If

			if cfg2 = true Then
			bstvalue
			DMDFLush
			DMD "cfg.png", "Saving", "Ball Save Time",  999999
			vpmtimer.addtimer 2000, "configoptions '"
			subconfig = False
			bst0 = False
			bst5 = False 
			bst10 = False
			bst15 = False 
			end If

			if cfg3 = true Then
			freevalue
			DMDFLush
			DMD "cfg.png", "Saving", "Free Play Setting",  999999
			vpmtimer.addtimer 2000, "configoptions '"
			subconfig = False
			gmcurrent = False
			gmcurrent = False 
			end If

			if cfg4 = true Then
			hsreset
			DMDFLush
			If hscurrent = True Then
			DMD "cfg.png", "High Scores", "Scores Reset",  999999
			Else
			DMD "cfg.png", "High Scores", "Scores Kept",  999999
			End If
			vpmtimer.addtimer 2000, "configoptions '"
			subconfig = False
			hscurrent = False
			end If

			if cfg5 = true Then
			profanityvalue
			DMDFLush
			If profanity = True Then
			DMD "cfg.png", "Profanity", "Kept",  999999
			Else
			DMD "cfg.png", "Profanity", "Removed",  999999
			End If
			vpmtimer.addtimer 2000, "configoptions '"
			subconfig = False
			profanity = False
			end If

			if cfg6 = true Then
			speedvalue
			DMDFLush
			DMD "cfg.png", "Saving", "Table Speed",  999999
			vpmtimer.addtimer 2000, "configoptions '"
			subconfig = False
			spd1 = False
			spd2 = False 
			spd3 = False
			end If
		Else
			if cfg1 = true Then
				cfgballs
			End If
			if cfg2 = true Then
				cfgbst
			End If
			if cfg3 = true Then
				cfgfree
			End If
			if cfg4 = true Then
				cfghs
			End If
			if cfg5 = true Then
				cfgprofanity
			End If
			if cfg6 = true Then
				cfgspeed
			End If
		End If
    end if
End Sub

Sub configexit
	inconfig = false
	DMDFLush
	StartAttractMode()
End Sub

Dim mainconfig
mainconfig = False

Sub configoptions
	DMDFLush
	mainconfig = True
	configmode = configmode + 1
	Select case configmode
		Case 0
		DMD "cfg.png", "0", "",  999999
		Case 1
		DMD "cfg.png", "Balls Per Game", bpgcurrent,  999999
		cfg1 = true
		cfg2 = False
		cfg3 = False
		cfg4 = False
		cfg5 = False
		cfg6 = False
		Case 2
		DMD "cfg.png", "Ball Save Time", bstcurrent,  999999
		cfg2 = True
		cfg1 = False
		cfg3 = False
		cfg4 = False
		cfg5 = False
		cfg6 = False
		Case 3
		DMD "cfg.png", "Free Play", gamemodecurrent,  999999
		cfg2 = False
		cfg1 = False
		cfg3 = True
		cfg4 = False
		cfg5 = False
		cfg6 = False
		Case 4
		DMD "cfg.png", "Reset High Scores?", "hs1=" &HighScore(0),  999999
		cfg2 = False
		cfg1 = False
		cfg3 = False
		cfg4 = True
		cfg5 = False
		cfg6 = False
		Case 5
		DMD "cfg.png", "Profanity", profanitycurrent,  999999
		cfg2 = False
		cfg1 = False
		cfg3 = False
		cfg4 = False
		cfg5 = True
		cfg6 = False
		Case 6
		DMD "cfg.png", "Table Speed", speedcurrent,  999999
		cfg2 = False
		cfg1 = False
		cfg3 = False
		cfg4 = False
		cfg5 = False
		cfg6 = True
		configmode = 0
	End Select
End Sub

'******
' cfg modes
'*******


' ****
' let's setup our defaults
' ****
Dim subconfig 'used to let the system know we're in a sub menu

Sub Loadbpg
' balls per game default set if empty
    bpgcurrent = LoadValue(TableName, "ballspergame")
    If bpgcurrent = "" Then SaveValue TableName, "ballspergame", 3 End If
' ball save time default set if empty
    bstcurrent = LoadValue(TableName, "ballsavetime")
    If bstcurrent = "" Then SaveValue TableName, "ballsavetime", 15 End If
' Game mode default set if empty
    gamemodecurrent = LoadValue(TableName, "gamemode")
    If gamemodecurrent = "" Then SaveValue TableName, "gamemode", "True" End If
' Profanity default set if empty
    profanitycurrent = LoadValue(TableName, "profanity")
    If profanitycurrent = "" Then SaveValue TableName, "profanity", "True" End If
' Speed default set if empty
    speedcurrent = LoadValue(TableName, "Speed")
    If speedcurrent = "" Then SaveValue TableName, "Speed", "Normal" End If
End Sub



' CFG 1 - balls per game
Dim cfg1
cfg1 = False
Dim cfg1ops
Dim bpg3
bpg3 = False
Dim bpg5
bpg5 = False
Dim bpgcurrent


Sub cfgballs
DMDFLush
mainconfig = False
subconfig = True
cfg1ops = cfg1ops + 1
Select case cfg1ops
	Case 1
		DMD "cfg.png", "Balls Per Game", "< 3 >",  999999
		bpg3 = True
		bpg5 = False
	Case 2
		DMD "cfg.png", "Balls Per Game", "< 5 >",  999999
		bpg5 = True
		bpg3 = False
		cfg1ops = 0
	End Select
End Sub

Sub bpgvalue
	If bpg5 = True Then
    SaveValue TableName, "ballspergame", 5
	Loadbpg
	End If
	If bpg3 = True Then
    SaveValue TableName, "ballspergame", 3
	Loadbpg
	End If
End Sub




' CFG 2 - ball save time
Dim cfg2
cfg2 = False
Dim cfg2ops
cfg2ops = 0
Dim bst0
bst0 = False
Dim bst5
bst5 = False
Dim bst10
bst10 = False
Dim bst15
bst15 = False
Dim bstcurrent

Sub cfgbst
DMDFLush
mainconfig = False
subconfig = True
cfg2ops = cfg2ops + 1
Select case cfg2ops
	Case 1
		DMD "cfg.png", "Ball Save Time", "< 0 >",  999999
		bst0 = True
		bst5 = False
		bst10 = False
		bst15 = False
	Case 2
		DMD "cfg.png", "Ball Save Time", "< 5 >",  999999
		bst0 = False
		bst5 = True
		bst10 = False
		bst15 = False
	Case 3
		DMD "cfg.png", "Ball Save Time", "< 10 >",  999999
		bst0 = False
		bst5 = False
		bst10 = True
		bst15 = False
	Case 4
		DMD "cfg.png", "Ball Save Time", "< 15 >",  999999
		bst0 = False
		bst5 = False
		bst10 = False
		bst15 = True
		cfg2ops = 0
	End Select
End Sub

Sub bstvalue
	If bst0 = True Then
    SaveValue TableName, "ballsavetime", 0
	Loadbpg
	End If
	If bst5 = True Then
    SaveValue TableName, "ballsavetime", 5
	Loadbpg
	End If
	If bst10 = True Then
    SaveValue TableName, "ballsavetime", 10
	Loadbpg
	End If
	If bst15 = True Then
    SaveValue TableName, "ballsavetime", 15
	Loadbpg
	End If
End Sub


' cfg 3 free play or coins 
    ' freeplay or coins  bFreePlay = True 'we dont want coins
Dim cfg3
cfg3 = False
Dim cfg3ops
Dim gmcurrent
gmcurrent = True
Dim gamemodecurrent

Sub cfgfree
DMDFLush
mainconfig = False
subconfig = True
cfg3ops = cfg3ops + 1
Select case cfg3ops
	Case 1
		DMD "cfg.png", "Free Play", "< True >",  999999
		gmcurrent = True
	Case 2
		DMD "cfg.png", "Free Play", "< False >",  999999
		gmcurrent = False
		cfg3ops = 0
	End Select
End Sub

Sub freevalue
	If gmcurrent = True Then
    SaveValue TableName, "gamemode", "True"
	Loadbpg
	End If
	If gmcurrent = False Then
    SaveValue TableName, "gamemode", "False"
	Loadbpg
	End If
End Sub

' cfg 4 reset high scores
Dim cfg4
cfg4 = False
Dim cfg4ops
Dim hscurrent
hscurrent = True

Sub cfghs
DMDFLush
mainconfig = False
subconfig = True
cfg4ops = cfg4ops + 1
Select case cfg4ops
	Case 1
		DMD "cfg.png", "Reset High Scores", "< Yes >",  999999
		hscurrent = True
	Case 2
		DMD "cfg.png", "Reset High Scores", "< Hell No >",  999999
		hscurrent = False
		cfg4ops = 0
	End Select
End Sub

Sub hsreset
	If hscurrent = True Then
    SaveValue TableName, "HighScore1", 11111111
    SaveValue TableName, "HighScore1Name", "011"
    SaveValue TableName, "HighScore2", 10000000
    SaveValue TableName, "HighScore2Name", "SBW"
    SaveValue TableName, "HighScore3", 9000000
    SaveValue TableName, "HighScore3Name", "JPS"
    SaveValue TableName, "HighScore4", 1000
    SaveValue TableName, "HighScore4Name", "BRB"
	Loadhs
	End If
	If hscurrent = False Then
	Loadhs
	End If
End Sub

' cfg 5 Profanity Filter
Dim cfg5
cfg5 = False
Dim cfg5ops
Dim profanity
profanity = True
Dim profanitycurrent

Sub cfgprofanity
DMDFLush
mainconfig = False
subconfig = True
cfg5ops = cfg5ops + 1
Select case cfg5ops
	Case 1
		DMD "cfg.png", "Profanity", "< True >",  999999
		profanity = True
	Case 2
		DMD "cfg.png", "Profanity", "< False >",  999999
		profanity = False
		cfg5ops = 0
	End Select
End Sub

Sub profanityvalue
	If profanity = True Then
    SaveValue TableName, "profanity", "True"
	Loadbpg
	End If
	If profanity = False Then
    SaveValue TableName, "profanity", "False"
	Loadbpg
	End If
End Sub

' CFG 6 - table speed
Dim cfg6
cfg6 = False
Dim cfg6ops
cfg6ops = 0
Dim spd1
spd1 = False
Dim spd2
spd2 = False
Dim spd3
spd3 = False
Dim speedcurrent

Sub cfgspeed
DMDFLush
mainconfig = False
subconfig = True
cfg6ops = cfg6ops + 1
Select case cfg6ops
	Case 1
		DMD "cfg.png", "Table Speed", "< Slow >",  999999
		spd1 = True
		spd2 = False
		spd3 = False
	Case 2
		DMD "cfg.png", "Table Speed", "< Normal >",  999999
		spd1 = False
		spd2 = True
		spd3 = False
	Case 3
		DMD "cfg.png", "Table Speed", "< Fast >",  999999
		spd1 = False
		spd2 = False
		spd3 = True
		cfg6ops = 0
	End Select
End Sub

Sub speedvalue
	If spd1 = True Then
    SaveValue TableName, "speed", "Slow"
	Loadbpg
	End If
	If spd2 = True Then
    SaveValue TableName, "speed", "Normal"
	Loadbpg
	End If
	If spd3 = True Then
    SaveValue TableName, "speed", "Fast"
	Loadbpg
	End If
End Sub

'*************
' Pause Table
'*************

Sub table1_Paused
End Sub

Sub table1_unPaused
End Sub

Sub table1_Exit
    Savehs
	If B2SOn Then Controller.Stop
End Sub

'********************
'     Flippers
'********************


Sub SolLFlipper(Enabled)
	If finalflips = False Then
	If lowerflippersoff = True Then
    If Enabled Then
        PlaySound SoundFXDOF("fx_flipperup", 101, DOFOn, DOFFlippers), 0, 1, -0.05, 0.15
        LeftFlipper.RotateToEnd
        If bSkillshotReady = False Then
            RotateLaneLightsLeft
        End If
    Else
        PlaySound SoundFXDOF("fx_flipperdown", 101, DOFOff, DOFFlippers), 0, 1, -0.05, 0.15
        LeftFlipper.RotateToStart
    End If
	End If
	End If
End Sub

Dim lowerflippersoff

Sub SolULFlipper(Enabled)
	If finalflips = False Then
	If lowerflippersoff = False Then
    If Enabled Then
        PlaySound SoundFXDOF("fx_flipperup", 101, DOFOn, DOFFlippers), 0, 1, -0.05, 0.15
        Flipper2.RotateToEnd
    Else
        PlaySound SoundFXDOF("fx_flipperdown", 101, DOFOff, DOFFlippers), 0, 1, -0.05, 0.15
        Flipper2.RotateToStart
    End If
	End If
	End If
End Sub

Sub SolRFlipper(Enabled)
	If finalflips = False Then
	If lowerflippersoff = True Then
    If Enabled Then
        PlaySound SoundFXDOF("fx_flipperup", 102, DOFOn, DOFFlippers), 0, 1, 0.05, 0.15
        RightFlipper.RotateToEnd
        If bSkillshotReady = False Then
            RotateLaneLightsRight
        End If
    Else
        PlaySound SoundFXDOF("fx_flipperdown", 102, DOFOff, DOFFlippers), 0, 1, 0.05, 0.15
        RightFlipper.RotateToStart
    End If
	End If
	End If
End Sub

Sub SolURFlipper(Enabled)
	If finalflips = False Then
	If lowerflippersoff = False Then
    If Enabled Then
        PlaySound SoundFXDOF("fx_flipperup", 101, DOFOn, DOFFlippers), 0, 1, -0.05, 0.15
        Flipper1.RotateToEnd
    Else
        PlaySound SoundFXDOF("fx_flipperdown", 101, DOFOff, DOFFlippers), 0, 1, -0.05, 0.15
        Flipper1.RotateToStart
    End If
	End If
	End If
End Sub

' flippers hit Sound

Sub LeftFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 10, -0.05, 0.25
End Sub

Sub RightFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 10, 0.05, 0.25
End Sub

Sub Flipper2_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 10, 0.05, 0.25
End Sub

Sub Flipper1_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 10, 0.05, 0.25
End Sub

Sub RotateLaneLightsLeft
    Dim TempState
    'flipper lanes
    TempState = ll1.State
    ll1.State = ll2.State
    ll2.State = ll3.State
    ll3.State = ll4.State
    ll4.State = ll5.State
    ll5.State = TempState
End Sub

Sub RotateLaneLightsRight
    Dim TempState
    'flipperlanes
    TempState = ll5.State
    ll5.State = ll4.State
    ll4.State = ll3.State
    ll3.State = ll2.State
    ll2.State = ll1.State
    ll1.State = TempState
End Sub


'*********
' TILT
'*********

'NOTE: The TiltDecreaseTimer Subtracts .01 from the "Tilt" variable every round

Sub CheckTilt                                    'Called when table is nudged
    Tilt = Tilt + TiltSensitivity                'Add to tilt count
    TiltDecreaseTimer.Enabled = True
    If(Tilt> TiltSensitivity) AND(Tilt <15) Then 'show a warning
        DMDFlush
        DMD "black.png", "CAREFUL!", "MOUTHBREATHER",  800
		DOF 131, DOFPulse
    End if
    If Tilt> 15 Then 'If more that 15 then TILT the table
        Tilted = True
        'display Tilt
        DMDFlush
        DMD "black.png", " ", "TILT!",  99999
        DisableTable True
        TiltRecoveryTimer.Enabled = True 'start the Tilt delay to check for all the balls to be drained
    End If
End Sub

Sub TiltDecreaseTimer_Timer
    ' DecreaseTilt
    If Tilt> 0 Then
        Tilt = Tilt - 0.1
    Else
        TiltDecreaseTimer.Enabled = False
    End If
End Sub

Sub DisableTable(Enabled)
    If Enabled Then
        'turn off GI and turn off all the lights
        GiOff
        LightSeqTilt.Play SeqAllOff
        'Disable slings, bumpers etc
        LeftFlipper.RotateToStart
        RightFlipper.RotateToStart
        'Bumper1.Force = 0

        LeftSlingshot.Disabled = 1
        RightSlingshot.Disabled = 1
    Else
        'turn back on GI and the lights
        GiOn
        LightSeqTilt.StopPlay
        'Bumper1.Force = 6
        LeftSlingshot.Disabled = 0
        RightSlingshot.Disabled = 0
        'clean up the buffer display
        DMDFlush
    End If
End Sub

Sub TiltRecoveryTimer_Timer()
    ' if all the balls have been drained then..
    If(BallsOnPlayfield = 0) Then
        ' do the normal end of ball thing (this doesn't give a bonus if the table is tilted)
        EndOfBall()
        TiltRecoveryTimer.Enabled = False
    End If
' else retry (checks again in another second or so)
End Sub

'********************
' Music as wav sounds
'********************

Dim Song
Song = ""

Sub PlaySong(name)
    If bMusicOn Then
        If Song <> name Then
            StopSound Song
            Song = name
            If Song = "m_end" Then
                PlaySound Song, 0, 0.1  'this last number is the volume, from 0 to 1
            Else
                PlaySound Song, -1, 0.1 'this last number is the volume, from 0 to 1
            End If
        End If
    End If
End Sub

'**********************
'     GI effects
' independent routine
' it turns on the gi
' when there is a ball
' in play
'**********************

Dim OldGiState
OldGiState = -1   'start witht the Gi off

Sub ChangeGi(col) 'changes the gi color
    Dim bulb
    For each bulb in aGILights
        SetLightColor bulb, col, -1
    Next
End Sub

Sub GIUpdateTimer_Timer
    Dim tmp, obj
    tmp = Getballs
    If UBound(tmp) <> OldGiState Then
        OldGiState = Ubound(tmp)
        If UBound(tmp) = 3 Then 'we have 4 captive balls on the table (-1 means no balls, 0 is the first ball, 1 is the second..)
            'GiOff               ' turn off the gi if no active balls on the table, we could also have used the variable ballsonplayfield.
        Else
            'Gion
        End If
    End If
End Sub

Sub GiOn
    DOF 126, DOFOn
    Dim bulb
    For each bulb in aGiLights
        bulb.State = 1
    Next

End Sub

Sub GiLowerOn
	Flasher1.opacity = 0
    'DOF 126, DOFOn
    Dim bulb
    For each bulb in lowergi
        bulb.State = 1
    Next

End Sub

Sub GiLowerOff
	Flasher1.opacity = 100
    'DOF 126, DOFOn
    Dim bulb
    For each bulb in lowergi
        bulb.State = 0
    Next

End Sub

Sub GiOff
    DOF 126, DOFOff
    Dim bulb
    For each bulb in aGiLights
        bulb.State = 0
    Next

End Sub

' GI & light sequence effects

Sub GiEffect(n)
    Select Case n
        Case 0 'all off
            LightSeqGi.Play SeqAlloff
        Case 1 'all blink
            LightSeqGi.UpdateInterval = 4
            LightSeqGi.Play SeqBlinking, , 5, 100
        Case 2 'random
            LightSeqGi.UpdateInterval = 10
            LightSeqGi.Play SeqRandom, 5, , 1000
        Case 3 'upon
            LightSeqGi.UpdateInterval = 4
            LightSeqGi.Play SeqUpOn, 5, 1
        Case 4 ' left-right-left
            LightSeqGi.UpdateInterval = 5
            LightSeqGi.Play SeqLeftOn, 10, 1
            LightSeqGi.UpdateInterval = 5
            LightSeqGi.Play SeqRightOn, 10, 1
    End Select
End Sub

Sub LightEffect(n)
    Select Case n
        Case 0 ' all off
            LightSeqInserts.Play SeqAlloff
        Case 1 'all blink
            LightSeqInserts.UpdateInterval = 4
            LightSeqInserts.Play SeqBlinking, , 5, 100
        Case 2 'random
            LightSeqInserts.UpdateInterval = 10
            LightSeqInserts.Play SeqRandom, 5, , 1000
        Case 3 'upon
            LightSeqInserts.UpdateInterval = 4
            LightSeqInserts.Play SeqUpOn, 10, 1
        Case 4 ' left-right-left
            LightSeqInserts.UpdateInterval = 5
            LightSeqInserts.Play SeqLeftOn, 10, 1
            LightSeqInserts.UpdateInterval = 5
            LightSeqInserts.Play SeqRightOn, 10, 1
        Case 5 'random
            LightSeqbumper.UpdateInterval = 4
            LightSeqbumper.Play SeqBlinking, , 5, 10
        Case 6 'random
            LightSeqRSling.UpdateInterval = 4
            LightSeqRSling.Play SeqBlinking, , 5, 6
        Case 7 'random
            LightSeqLSling.UpdateInterval = 4
            LightSeqLSling.Play SeqBlinking, , 5, 6
        Case 8 'random
            LightSeqBack.UpdateInterval = 4
            LightSeqBack.Play SeqBlinking, , 5, 6
        Case 9 'random
            LightSeqTruck.UpdateInterval = 4
            LightSeqTruck.Play SeqBlinking, , 5, 10
        Case 10 'random
            LightSeqbarb.UpdateInterval = 4
            LightSeqbarb.Play SeqBlinking, , 5, 10
        Case 11 'random
            LightSeqdice.UpdateInterval = 4
            LightSeqdice.Play SeqBlinking, , 5, 10
        Case 12 'random
            LightSeqlr.UpdateInterval = 4
            LightSeqlr.Play SeqBlinking, , 5, 10
        Case 13 'random
            LightSeqrr.UpdateInterval = 4
            LightSeqrr.Play SeqBlinking, , 5, 10
    End Select
End Sub

' Flasher Effects using lights

Dim FEStep, FEffect
FEStep = 0
FEffect = 0

Sub FlashEffect(n)
    Select case n
        Case 0 ' all off
            LightSeqFlasher.Play SeqAlloff
        Case 1 'all blink
            LightSeqFlasher.UpdateInterval = 4
            LightSeqFlasher.Play SeqBlinking, , 5, 100
        Case 2 'random
            LightSeqFlasher.UpdateInterval = 10
            LightSeqFlasher.Play SeqRandom, 5, , 1000
        Case 3 'upon
            LightSeqFlasher.UpdateInterval = 4
            LightSeqFlasher.Play SeqUpOn, 10, 1
        Case 4 ' left-right-left
            LightSeqFlasher.UpdateInterval = 5
            LightSeqFlasher.Play SeqLeftOn, 10, 1
            LightSeqFlasher.UpdateInterval = 5
            LightSeqFlasher.Play SeqRightOn, 10, 1
        Case 5 ' top flashers blink fast
            FlashForms f5, 1000, 50, 0
            FlashForms f6, 1000, 50, 0
            FlashForms f7, 1000, 50, 0
            FlashForms f8, 1000, 50, 0
            FlashForms f9, 1000, 50, 0
    End Select
End Sub


lrflashtime.enabled = 0
Dim letsflash
Sub lrflashnow
	lrflashtime.enabled = 1
End Sub
Sub lrflashtime_Timer 
	letsflash = letsflash + 1
	Select Case letsflash
	Case 0
		FlasherLeftRed.opacity = 0
	Case 1
		FlasherLeftRed.opacity = 100
	Case 2
		FlasherLeftRed.opacity = 0
	Case 3
		FlasherLeftRed.opacity = 100
	Case 4
		FlasherLeftRed.opacity = 0
	Case 5
		FlasherLeftRed.opacity = 100
	Case 6
		FlasherLeftRed.opacity = 0
		lrflashtime.Enabled = False
		letsflash = 0
	End Select
End Sub


rrflashtime.enabled = 0
Dim letsflash2
Sub rrflashnow
	rrflashtime.enabled = 1
End Sub
Sub rrflashtime_Timer 
	letsflash2 = letsflash2 + 1
	Select Case letsflash2
	Case 0
		FlasherRightWhite.opacity = 0
	Case 1
		FlasherRightWhite.opacity = 100
	Case 2
		FlasherRightWhite.opacity = 0
	Case 3
		FlasherRightWhite.opacity = 100
	Case 4
		FlasherRightWhite.opacity = 0
	Case 5
		FlasherRightWhite.opacity = 100
	Case 6
		FlasherRightWhite.opacity = 0
		rrflashtime.Enabled = False
		letsflash2 = 0
	End Select
End Sub

Sub Flashxmas(n)
    Select case n
        Case 0 ' all off
            LightSeqaxmas.Play SeqAlloff
        Case 1 'all blink
            LightSeqaxmas.UpdateInterval = 4
            LightSeqaxmas.Play SeqBlinking, , 5, 100
        Case 2 'random
            LightSeqaxmas.UpdateInterval = 10
            LightSeqaxmas.Play SeqRandom, 5, , 1000
    End Select
End Sub

' *********************************************************************
'                      Supporting Ball & Sound Functions
' *********************************************************************

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
    Vol = Csng(BallVel(ball) ^2 / 1500)
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "table1" is the name of the table
    Dim tmp
    tmp = ball.x * 2 / table1.width-1
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
    BallVel = INT(SQR((ball.VelX ^2) + (ball.VelY ^2) ) )
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
    If UBound(BOT) = -1 Then Exit Sub 'there no extra balls on this table

    ' play the rolling sound for each ball
    For b = lob to UBound(BOT)
        If BallVel(BOT(b) )> 1 Then
            If BOT(b).z <30 Then
                ballpitch = Pitch(BOT(b) )
            Else
                ballpitch = Pitch(BOT(b) ) * 50
            End If
            rolling(b) = True
            PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) ), Pan(BOT(b) ), 0, ballpitch, 1, 0
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

'******************************
' Diverse Collection Hit Sounds
'******************************

Sub aMetals_Hit(idx):PlaySound "fx_MetalHit", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
Sub aRubber_Bands_Hit(idx):PlaySound "fx_rubber_band", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
Sub aRubber_Posts_Hit(idx):PlaySound "fx_postrubber", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
Sub aRubber_Pins_Hit(idx):PlaySound "fx_rubber", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
Sub aPlastics_Hit(idx):PlaySound "fx_PlasticHit", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
Sub aGates_Hit(idx):PlaySound "fx_Gate", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
Sub aWoods_Hit(idx):PlaySound "fx_Woodhit", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub

' Random quotes from the game

Sub PlayQuote_timer() 'one quote each 2 minutes
    Dim Quote
    Quote = "di_quote" & INT(RND * 56) + 1
    PlaySound Quote
End Sub

' Ramp Soundss
Sub REnd1_Hit()
    StopSound "fx_metalrolling"
    PlaySound "fx_ballrampdrop", 0, 1, pan(ActiveBall)
End Sub

Sub REnd2_Hit()
    StopSound "fx_metalrolling"
    PlaySound "fx_ballrampdrop", 0, 1, pan(ActiveBall)
End Sub

' *********************************************************************
'                        User Defined Script Events
' *********************************************************************

' Initialise the Table for a new Game
'
Sub ResetForNewGame()
    Dim i

    bGameInPLay = True

    'resets the score display, and turn off attrack mode
    StopAttractMode
    GiOn

    TotalGamesPlayed = TotalGamesPlayed + 1
    CurrentPlayer = 1
    PlayersPlayingGame = 1
    bOnTheFirstBall = True
    For i = 1 To MaxPlayers
        Score(i) = 0
        BonusPoints(i) = 0
        BonusHeldPoints(i) = 0
        BonusMultiplier(i) = 1
        BallsRemaining(i) = bpgcurrent
        ExtraBallsAwards(i) = 0
    Next
	'For x = 0 to UBOUND(aLights)
	'	lamps(CurrentPlayer,x) = x.State
	'Next
    ' initialise any other flags
    Tilt = 0

    ' initialise Game variables
    Game_Init()

    ' you may wish to start some music, play a sound, do whatever at this point

    ' set up the start delay to handle any Start of Game Attract Sequence
    vpmtimer.addtimer 1500, "FirstBall '"
End Sub

' This is used to delay the start of a game to allow any attract sequence to
' complete.  When it expires it creates a ball for the player to start playing with

Sub FirstBall
    ' reset the table for a new ball
    ResetForNewPlayerBall()
    ' create a new ball in the shooters lane
    CreateNewBall()
End Sub

' (Re-)Initialise the Table for a new ball (either a new ball after the player has
' lost one or we have moved onto the next player (if multiple are playing))

Sub ResetForNewPlayerBall()
    ' make sure the correct display is upto date
    AddScore 0

    ' reset any drop targets, lights, game Mode etc..

'	For x = 0 to UBOUND(aLights)
'		x.State = lamps(CurrentPlayer,x)
'	Next

    BonusPoints(CurrentPlayer) = 0
    bBonusHeld = False
    bExtraBallWonThisBall = False
    ResetNewBallLights()
    'Reset any table specific
    ResetNewBallVariables

    'This is a new ball, so activate the ballsaver
    bBallSaverReady = True
	RaiseTargets

    'and the skillshot
    bSkillShotReady = True

'Change the music ?
End Sub

' Create a new ball on the Playfield

Sub CreateNewBall()
    ' create a ball in the plunger lane kicker.
    BallRelease.CreateSizedball BallSize / 2

    ' There is a (or another) ball on the playfield
    BallsOnPlayfield = BallsOnPlayfield + 1

    ' kick it out..
	PlaySound SoundFXDOF("fx_Ballrel", 114, DOFPulse, DOFContactors), 0, 1, 0.1, 0.1
    BallRelease.Kick 90, 4

' if there is 2 or more balls then set the multibal flag (remember to check for locked balls and other balls used for animations)
' set the bAutoPlunger flag to kick the ball in play automatically
    If BallsOnPlayfield > 1 Then
        bMultiBallMode = True
			DOF 127, DOFPulse
        bAutoPlunger = True
    End If

	If barbMultiball = True Then
        PlaySong "m_barb"  'this last number is the volume, from 0 to 1
	End If

	If runMultiball = True Then
        PlaySong "m_run"  'this last number is the volume, from 0 to 1
	End If

	If willMultiball = True Then
        PlaySong "m_multiball"  'this last number is the volume, from 0 to 1
	End If

	If demoMultiball = True Then
        PlaySong "m_multiball"  'this last number is the volume, from 0 to 1
	End If
End Sub

' Add extra balls to the table with autoplunger
' Use it as AddMultiball 4 to add 4 extra balls to the table

Sub AddMultiball(nballs)
    mBalls2Eject = mBalls2Eject + nballs
    CreateMultiballTimer.Enabled = True
End Sub

' Eject the ball after the delay, AddMultiballDelay
Sub CreateMultiballTimer_Timer()
    ' wait if there is a ball in the plunger lane
    If bBallInPlungerLane Then
        Exit Sub
    Else
        If BallsOnPlayfield <MaxMultiballs Then
            CreateNewBall()
            mBalls2Eject = mBalls2Eject -1
            If mBalls2Eject = 0 Then 'if there are no more balls to eject then stop the timer
                Me.Enabled = False
            End If
        Else 'the max number of multiballs is reached, so stop the timer
            mBalls2Eject = 0
            Me.Enabled = False
        End If
    End If
End Sub

' The Player has lost his ball (there are no more balls on the playfield).
' Handle any bonus points awarded

Sub EndOfBall()
    'Dim AwardPoints, TotalBonus, ii
    'AwardPoints = 0
    'TotalBonus = 0
    ' the first ball has been lost. From this point on no new players can join in
    bOnTheFirstBall = False

    ' only process any of this if the table is not tilted.  (the tilt recovery
    ' mechanism will handle any extra balls or end of game)

    If NOT Tilted Then

        ' add a bit of a delay to allow for the bonus points to be shown & added up
        vpmtimer.addtimer 100, "EndOfBall2 '"
    Else 'if tilted then only add a short delay
        vpmtimer.addtimer 100, "EndOfBall2 '"
    End If
End Sub

' The Timer which delays the machine to allow any bonus points to be added up
' has expired.  Check to see if there are any extra balls for this player.
' if not, then check to see if this was the last ball (of the currentplayer)
'
Sub EndOfBall2()
    ' if were tilted, reset the internal tilted flag (this will also
    ' set TiltWarnings back to zero) which is useful if we are changing player LOL
    Tilted = False
    Tilt = 0
    DisableTable False 'enable again bumpers and slingshots

    ' has the player won an extra-ball ? (might be multiple outstanding)
    If(ExtraBallsAwards(CurrentPlayer) <> 0) Then
        'debug.print "Extra Ball"

        ' yep got to give it to them
        ExtraBallsAwards(CurrentPlayer) = ExtraBallsAwards(CurrentPlayer) - 1

        ' if no more EB's then turn off any shoot again light
        If(ExtraBallsAwards(CurrentPlayer) = 0) Then
            LightShootAgain.State = 0
        End If

        ' You may wish to do a bit of a song AND dance at this point
				LightSeqFlasher.UpdateInterval = 150
				LightSeqFlasher.Play SeqRandom, 10, , 2000
				DMD "black.png", "SHOOT", "AGAIN",  2000

        ' Create a new ball in the shooters lane
        CreateNewBall()
    Else ' no extra balls

        BallsRemaining(CurrentPlayer) = BallsRemaining(CurrentPlayer) - 1

        ' was that the last ball ?
        If(BallsRemaining(CurrentPlayer) <= 0) Then
            'debug.print "No More Balls, High Score Entry"

            ' Submit the currentplayers score to the High Score system
            CheckHighScore()
        ' you may wish to play some music at this point

        Else

            ' not the last ball (for that player)
            ' if multiple players are playing then move onto the next one
            EndOfBallComplete()
        End If
    End If
End Sub

' This function is called when the end of bonus display
' (or high score entry finished) AND it either end the game or
' move onto the next player (or the next ball of the same player)
'
Sub EndOfBallComplete()
    Dim NextPlayer

    'debug.print "EndOfBall - Complete"

    ' are there multiple players playing this game ?
    If(PlayersPlayingGame> 1) Then
        ' then move to the next player
        NextPlayer = CurrentPlayer + 1
        ' are we going from the last player back to the first
        ' (ie say from player 4 back to player 1)
        If(NextPlayer> PlayersPlayingGame) Then
            NextPlayer = 1
        End If
    Else
        NextPlayer = CurrentPlayer
    End If

    'debug.print "Next Player = " & NextPlayer

    ' is it the end of the game ? (all balls been lost for all players)
    If((BallsRemaining(CurrentPlayer) <= 0) AND(BallsRemaining(NextPlayer) <= 0) ) Then
        ' you may wish to do some sort of Point Match free game award here
        ' generally only done when not in free play mode

        ' set the machine into game over mode
        EndOfGame()

    ' you may wish to put a Game Over message on the desktop/backglass

    Else
        ' set the next player
        CurrentPlayer = NextPlayer

        ' make sure the correct display is up to date
        AddScore 0

        ' reset the playfield for the new player (or new ball)
        ResetForNewPlayerBall()

        ' AND create a new ball
        CreateNewBall()

        ' play a sound if more than 1 player
        If PlayersPlayingGame> 1 Then
            PlaySound "vo_player" &CurrentPlayer
            DMD "black.png", " ", "PLAYER " &CurrentPlayer,  800
        End If
    End If
End Sub

' This function is called at the End of the Game, it should reset all
' Drop targets, AND eject any 'held' balls, start any attract sequences etc..

Sub EndOfGame()
    'debug.print "End Of Game"
	introposition = 0
    bGameInPLay = False
    ' just ended your game then play the end of game tune
    If NOT bJustStarted Then
        PlaySong "m_end"
    End If
    bJustStarted = False
    ' ensure that the flippers are down
	SolLFlipper 0
	SolRFlipper 0
	SolULFlipper 0
	SolURFlipper 0
	BallsInLock(CurrentPlayer) = 0
	BallsInRunLock(CurrentPlayer) = 0
	RaiseTargets

    ' terminate all Mode - eject locked balls
    ' most of the Mode/timers terminate at the end of the ball
    'PlayQuote.Enabled = 0
    ' show game over on the DMD
	DMD "black.png", "Game Over", "",  2000
	Dim i
	If Score(1) Then
		DMD "black.png", "PLAYER 1", Score(1), 3000
	End If
	If Score(2) Then
		DMD "black.png", "PLAYER 2", Score(2), 3000
	End If
	If Score(3) Then
		DMD "black.png", "PLAYER 3", Score(3), 3000
	End If
	If Score(4) Then
		DMD "black.png", "PLAYER 4", Score(4), 3000
	End If

    ' set any lights for the attract mode
    GiOff
    StartAttractMode
' you may wish to light any Game Over Light you may have
End Sub

Function Balls
    Dim tmp
    tmp = bpgcurrent - BallsRemaining(CurrentPlayer) + 1
    If tmp> bpgcurrent Then
        Balls = bpgcurrent
    Else
        Balls = tmp
    End If
End Function

' *********************************************************************
'                      Drain / Plunger Functions
' *********************************************************************

' lost a ball ;-( check to see how many balls are on the playfield.
' if only one then decrement the remaining count AND test for End of game
' if more than 1 ball (multi-ball) then kill of the ball but don't create
' a new one
'
Dim Saves
Dim Drains
Sub Ballsaved
	Saves = Saves + 1
	Select Case Saves
        Case 1 DMD "ballsave1-2.wmv", "", "", 3000
		Case 2 DMD "ballsave2-6.wmv", "", "", 7000:Saves = 0
	End Select
End Sub
Sub Balldrained
	If profanitycurrent = "True" Then
	Drains = Drains + 1
	Select Case Drains
        Case 1 DMD "drain3-3.wmv", "", "", 4000
		Case 2 DMD "drain2-2.wmv", "", "", 3000
		Case 3 DMD "drain1-4.wmv", "", "", 5000:Drains = 0
	End Select
	End If
	If profanitycurrent = "False" Then
	Drains = Drains + 1
	Select Case Drains
        Case 1 DMD "drain3-3.wmv", "", "", 4000
		Case 2 DMD "drain1-4.wmv", "", "", 5000:Drains = 0
	End Select
	End If
End Sub

Sub Drain_Hit()
    startB2S(7)
    ' Destroy the ball
    Drain.DestroyBall
    ' Exit Sub ' only for debugging
		BallsOnPlayfield = BallsOnPlayfield - 1

    ' pretend to knock the ball into the ball storage mech
    PlaySound "fx_drain"
    'if Tilted the end Ball Mode
    If Tilted Then
        StopEndOfBallMode
    End If

    ' if there is a game in progress AND it is not Tilted

    If(bGameInPLay = True) AND(Tilted = False) Then

		' is the ball saver active,
        If(bBallSaverActive = True) Then
				' yep, create a new ball in the shooters lane
				' we use the Addmultiball in case the multiballs are being ejected
				AddMultiball 1
				' we kick the ball with the autoplunger
				bAutoPlunger = True

				' you may wish to put something on a display or play a sound at this point
				If bMultiBallMode = False Then
				
				Ballsaved
				End If
        Else
            ' cancel any multiball if on last ball (ie. lost all other balls)
            If(BallsOnPlayfield = 1) Then
                ' AND in a multi-ball??
                If(bMultiBallMode = True) then
                    ' not in multiball mode any more
					CheckKIDSLane
					CheckGUARDTargets
					Gate7.Open = True
					If barbMultiball = True Then
					EndBarb
					End If
					If RunMultiball = True Then
					EndRun
					End If
					If WillMultiball = True Then
					EndWill
					End If
					If DemoMultiball = True Then
					EndDemo
					End If
                    bMultiBallMode = False		
                    ChangeGi "white"
                    ' you may wish to change any music over at this point and
                    ' turn off any multiball specific lights
                    'ResetJackpotLights
                    CurrentSong
                End If
				CurrentSong
				bMultiBallMode = False		
                ChangeGi "white"
				CheckKIDSLane
				CheckGUARDTargets
		End If

            ' was that the last ball on the playfield
            If(BallsOnPlayfield = 0) Then
                ' End Mode and timers
                PlaySong "m_wait"
                ChangeGi "white"
			' Show the end of ball animation and handle the end of ball (count bonus, change player, high score entry etc..)
			'  DMD "ball-lost-2.wmv", "", "",  4500
			Balldrained
            vpmtimer.addtimer 4500, "EndOfBall '"
                StopEndOfBallMode
            End If
        End If
    End If
End Sub

' The Ball has rolled out of the Plunger Lane and it is pressing down the trigger in the shooters lane
' Check to see if a ball saver mechanism is needed and if so fire it up.
Dim changetrack
Sub Whatsong
	changetrack = changetrack +1
End Sub

Sub CurrentSong
	Select Case changetrack
		Case 1
			PlaySong "m_main"
		Case 2
			PlaySong "m_main2"
		Case 3
			PlaySong "m_main3"
			changetrack = 0
		End Select
End Sub

Sub ballsavestarttrigger_hit
    ' if there is a need for a ball saver, then start off a timer		    ' if there is a need for a ball saver, then start off a timer
    ' if there is a need for a ball saver, then start off a timer
    ' only start if it is ready, and it is currently not running, else it will reset the time period
    If(bBallSaverReady = True) AND(bstcurrent <> 0) And(bBallSaverActive = False) Then
        EnableBallSaver bstcurrent
    End If
End Sub

Sub swPlungerRest_Hit()
    'debug.print "ball in plunger lane"
    ' some sound according to the ball position
    PlaySound "fx_sensor", 0, 1, 0.15, 0.25
    bBallInPlungerLane = True
    ' turn on Launch light is there is one
    'LaunchLight.State = 2
    'Update the Scoreboard
    DMDScoreNow
    ' kick the ball in play if the bAutoPlunger flag is on
    If bAutoPlunger Then
        'debug.print "autofire the ball"
        	PlungerIM.Strength = 45
			PlungerIM.AutoFire
			PlungerIM.Strength = Plunger.MechStrength
			DOF 114, DOFPulse		
		DOF 115, DOFPulse
        bAutoPlunger = False
    End If
	DOF 141, DOFOn
    'Start the Selection of the skillshot if ready
    If bSkillShotReady Then
        swPlungerRest.TimerEnabled = 1 ' this is a new ball, so show the launch ball if inactive for 6 seconds
        UpdateSkillshot()
    If NOT bMultiballMode Then
        Whatsong
		CurrentSong
    End If
    End If
    ' remember last trigger hit by the ball.
    LastSwitchHit = "swPlungerRest"
End Sub

' The ball is released from the plunger turn off some flags and check for skillshot

Sub swPlungerRest_UnHit()
    bBallInPlungerLane = False
    swPlungerRest.TimerEnabled = 0 'stop the launch ball timer if active
    DMDScoreNow
    If bSkillShotReady Then
        ResetSkillShotTimer.Enabled = 1
    End If
' turn off LaunchLight
'LaunchLight.State = 0
End Sub

' swPlungerRest timer to show the "launch ball" if the player has not shot the ball during 6 seconds

Sub swPlungerRest_Timer
    DMD "start-5.wmv", "", "",  6000
    swPlungerRest.TimerEnabled = 0
End Sub

Sub EnableBallSaver(seconds)
    'debug.print "Ballsaver started"
    ' set our game flag
    bBallSaverActive = True
    bBallSaverReady = False
    ' start the timer
    BallSaverTimerExpired.Interval = 1000 * seconds
    BallSaverTimerExpired.Enabled = True
    BallSaverSpeedUpTimer.Interval = 1000 * seconds -(1000 * seconds) / 3
    BallSaverSpeedUpTimer.Enabled = True
    ' if you have a ball saver light you might want to turn it on at this point (or make it flash)
    LightShootAgain.BlinkInterval = 160
    LightShootAgain.State = 2
End Sub

' The ball saver timer has expired.  Turn it off AND reset the game flag
'
Sub BallSaverTimerExpired_Timer()
    'debug.print "Ballsaver ended"
    BallSaverTimerExpired.Enabled = False
    ' clear the flag
    bBallSaverActive = False
    ' if you have a ball saver light then turn it off at this point
    LightShootAgain.State = 0
End Sub

Sub BallSaverSpeedUpTimer_Timer()
    'debug.print "Ballsaver Speed Up Light"
    BallSaverSpeedUpTimer.Enabled = False
    ' Speed up the blinking
    LightShootAgain.BlinkInterval = 80
    LightShootAgain.State = 2
End Sub

' *********************************************************************
'                      Supporting Score Functions
' *********************************************************************

' Add points to the score AND update the score board
'
Sub AddScore(points)
    If(Tilted = False) Then
        ' add the points to the current players score variable
        Score(CurrentPlayer) = Score(CurrentPlayer) + points
        ' update the score displays
        DMDScore
    End if

' you may wish to check to see if the player has gotten a replay
End Sub

' Add bonus to the bonuspoints AND update the score board

Sub AddBonus(points) 'not used in this table, since there are many different bonus items.
    If(Tilted = False) Then
        ' add the bonus to the current players bonus variable
        BonusPoints(CurrentPlayer) = BonusPoints(CurrentPlayer) + points
        ' update the score displays
        DMDScore
    End if

' you may wish to check to see if the player has gotten a replay
End Sub

' Add some points to the current Jackpot.
'
Sub AddJackpot(points) 'not used in this table
' Jackpots only generally increment in multiball mode AND not tilted
' but this doesn't have to be the case
	If(Tilted = False)Then
		If(bMultiBallMode = True) Then
		Jackpot = Jackpot + points
' you may wish to limit the jackpot to a upper limit, ie..
		If (Jackpot >= 6000) Then
		Jackpot = 6000
		End if
	End if
	End if
End Sub

Sub AddSuperJackpot(points)
    If(Tilted = False) Then

        ' If(bMultiBallMode = True) Then
        SuperJackpot = SuperJackpot + points
    ' you may wish to limit the jackpot to a upper limit, ie..
    '	If (Jackpot >= 6000) Then
    '		Jackpot = 6000
    ' 	End if
    'End if
    End if
End Sub

' Set the Bonus Multiplier to the specified level AND set any lights accordingly
' There is no bonus multiplier lights in this table

Sub AwardExtraBall()
    If NOT bExtraBallWonThisBall Then
		DMD "black.png", "EXTRA", "BALL",  2000
		LightShootAgain.State = 1
		LightSeqFlasher.UpdateInterval = 150
		LightSeqFlasher.Play SeqRandom, 10, , 10000
        ExtraBallsAwards(CurrentPlayer) = ExtraBallsAwards(CurrentPlayer) + 1
        bExtraBallWonThisBall = True
    Else
	AddScore 2000000
	END If
End Sub

Sub AwardSpecial()
    Credits = Credits + 1
    DOF 140, DOFOn
	PlaySound SoundFXDOF("knocker",136,DOFPulse,DOFKnocker)		
	DOF 115, DOFPulse
    GiEffect 1
    LightEffect 1
End Sub

'in this table the jackpot is always 1 million + 10% of your score

Sub AwardJackpot() 'award a normal jackpot, double or triple jackpot
    Jackpot = 1000000 + Round(Score(CurrentPlayer) / 10, 0)
    DMD "black.png", "JACKPOT", Jackpot,  1000
    AddScore Jackpot
	DOF 125, DOFPulse
    GiEffect 1
    LightEffect 2
    FlashEffect 2
End Sub

Sub AwardDoubleJackpot() 'in this table the jackpot is always 1 million + 10% of your score
    Jackpot = (1000000 + Round(Score(CurrentPlayer) / 10, 0) ) * 2
    DMD "black.png", "DOUBLE JACKPOT", Jackpot,  1000
	DOF 125, DOFPulse
    GiEffect 1
    LightEffect 2
    FlashEffect 2
End Sub

Sub AwardTripleJackpot() 'in this table the jackpot is always 1 million + 10% of your score
    'DOF 132, DOFPulse
    Jackpot = (1000000 + Round(Score(CurrentPlayer) / 10, 0) ) * 3
    DMD "black.png", "TRIPLE JACKPOT", Jackpot,  1000
	DOF 125, DOFPulse
    AddScore Jackpot
    GiEffect 1
    LightEffect 2
    FlashEffect 2
End Sub

Sub AwardSuperJackpot() 'in this table a super jackpot is a jackpot when the playfield multiplier is over 4x or 5x
    Dim tmp
    tmp = "vo_superjackpot" & INT(RND * 5 + 1)
    DMDFlush
    DMD "black.png", "SUPER JACKPOT", Jackpot,  1000
	DOF 125, DOFPulse
    AddScore SuperJackpot
    GiEffect 1
    LightEffect 2
    FlashEffect 2
End Sub

Sub AwardSkillshot()
    Dim i
	DOF 125, DOFPulse
    ResetSkillShotTimer_Timer
    'show dmd animation
    DMDFlush
	run1.IsDropped = 1
	run2.IsDropped = 1
	run3.IsDropped = 1
	lr1.State = 1
	lr2.State = 1
	lr3.State = 1
    DMD "steve6-2.wmv", "", "",  3000
    AddScore SkillshotValue(CurrentPLayer)
    SkillShotValue(CurrentPLayer) = SkillShotValue(CurrentPLayer) + 1000000
    'do some light show
	'CheckRunTargets
    GiEffect 1
    LightEffect 2
	LightEffect 9
End Sub

Sub Congratulation()
    Dim tmp
    tmp = "vo_congrat" & INT(RND * 21 + 1)
    PlaySound tmp
End Sub
'*****************************
'    Load / Save / Highscore
'*****************************

Sub Loadhs
    Dim x
    x = LoadValue(TableName, "HighScore1")
    If(x <> "") Then HighScore(0) = CDbl(x) Else HighScore(0) = 1300000 End If

    x = LoadValue(TableName, "HighScore1Name")
    If(x <> "") Then HighScoreName(0) = x Else HighScoreName(0) = "SBW" End If

    x = LoadValue(TableName, "HighScore2")
    If(x <> "") then HighScore(1) = CDbl(x) Else HighScore(1) = 1200000 End If

    x = LoadValue(TableName, "HighScore2Name")
    If(x <> "") then HighScoreName(1) = x Else HighScoreName(1) = "JPS" End If

    x = LoadValue(TableName, "HighScore3")
    If(x <> "") then HighScore(2) = CDbl(x) Else HighScore(2) = 1100000 End If

    x = LoadValue(TableName, "HighScore3Name")
    If(x <> "") then HighScoreName(2) = x Else HighScoreName(2) = "011" End If

    x = LoadValue(TableName, "HighScore4")
    If(x <> "") then HighScore(3) = CDbl(x) Else HighScore(3) = 1000000 End If

    x = LoadValue(TableName, "HighScore4Name")
    If(x <> "") then HighScoreName(3) = x Else HighScoreName(3) = "011" End If

    x = LoadValue(TableName, "Credits")
    If(x <> "") then Credits = CInt(x) Else Credits = 0 End If

    'x = LoadValue(TableName, "Jackpot")
    'If(x <> "") then Jackpot = CDbl(x) Else Jackpot = 200000 End If
    x = LoadValue(TableName, "TotalGamesPlayed")
    If(x <> "") then TotalGamesPlayed = CInt(x) Else TotalGamesPlayed = 0 End If
End Sub

Sub Savehs
    SaveValue TableName, "HighScore1", HighScore(0)
    SaveValue TableName, "HighScore1Name", HighScoreName(0)
    SaveValue TableName, "HighScore2", HighScore(1)
    SaveValue TableName, "HighScore2Name", HighScoreName(1)
    SaveValue TableName, "HighScore3", HighScore(2)
    SaveValue TableName, "HighScore3Name", HighScoreName(2)
    SaveValue TableName, "HighScore4", HighScore(3)
    SaveValue TableName, "HighScore4Name", HighScoreName(3)
    SaveValue TableName, "Credits", Credits
    'SaveValue TableName, "Jackpot", Jackpot
    SaveValue TableName, "TotalGamesPlayed", TotalGamesPlayed
End Sub

' ***********************************************************
'  High Score Initals Entry Functions - based on Black's code
' ***********************************************************

Dim hsbModeActive
Dim hsEnteredName
Dim hsEnteredDigits(3)
Dim hsCurrentDigit
Dim hsValidLetters
Dim hsCurrentLetter
Dim hsLetterFlash

Sub CheckHighscore()
    Dim tmp
    tmp = Score(1)

    If Score(2)> tmp Then tmp = Score(2)
    If Score(3)> tmp Then tmp = Score(3)
    If Score(4)> tmp Then tmp = Score(4)

    If tmp> HighScore(1) Then 'add 1 credit for beating the highscore
        AwardSpecial
    End If

    If tmp> HighScore(3) Then
        vpmtimer.addtimer 2000, "PlaySound ""vo_contratulationsgreatscore"" '"
        HighScore(3) = tmp
        'enter player's name
        HighScoreEntryInit()
    Else
        EndOfBallComplete()
    End If
End Sub

Sub HighScoreEntryInit()
    hsbModeActive = True
    PlaySound "vo_enteryourinitials"
    hsLetterFlash = 0

    hsEnteredDigits(0) = "A"
    hsEnteredDigits(1) = " "
    hsEnteredDigits(2) = " "
    hsCurrentDigit = 0

    hsValidLetters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ<" ' < is back arrow
    hsCurrentLetter = 1
    DMDFlush
    DMD "fw-6.wmv", "You Got", "A High Score",  3000
    DMD "black.png", "Your Score", Score(CurrentPlayer),  2000
    DMD "black.png", "Enter Your", "Initials",  250
	Playsound "bellhs"
    DMD "black.png", "", "",  250
    DMD "black.png", "Enter Your", "Initials",  250
	Playsound "bellhs"
    DMD "black.png", "", "",  250
    DMD "black.png", "Enter Your", "Initials",  250
	Playsound "bellhs"
    DMD "black.png", "", "",  250
    DMD "black.png", "Enter Your", "Initials",  250
	Playsound "bellhs"
    DMD "black.png", "", "",  250
    'DMD "highscore-20.wmv", "", "",  19000
    DMDId "hsc", "hsb.jpg", " Initials:A  ", " ",  999999
    vpmtimer.addtimer 8000, "HighScoreDisplayName() '"	
End Sub

Sub EnterHighScoreKey(keycode)
    If keycode = LeftFlipperKey Then
        Playsound "fx_Previous"
        hsCurrentLetter = hsCurrentLetter - 1
        hsLetter = hsLetter - 1
        if(hsCurrentLetter = 0) then
            hsCurrentLetter = len(hsValidLetters)
        end if
        HighScoreDisplayName()
    End If

    If keycode = RightFlipperKey Then
        Playsound "fx_Next"
        hsCurrentLetter = hsCurrentLetter + 1
        hsLetter = hsLetter + 1
        if(hsCurrentLetter> len(hsValidLetters) ) then
            hsCurrentLetter = 1
        end if
        HighScoreDisplayName()
    End If

    If keycode = StartGameKey or keycode = PlungerKey Then
        if(mid(hsValidLetters, hsCurrentLetter, 1) <> "<") then
            playsound "fx_Enter"
            hsEnteredDigits(hsCurrentDigit) = mid(hsValidLetters, hsCurrentLetter, 1)
            hsCurrentDigit = hsCurrentDigit + 1
            if(hsCurrentDigit = 3) then
                HighScoreCommitName()
            else
                HighScoreDisplayName()
            end if
        else
            playsound "fx_Esc"
            hsEnteredDigits(hsCurrentDigit) = " "
            if(hsCurrentDigit> 0) then
                hsCurrentDigit = hsCurrentDigit - 1
            end if
            HighScoreDisplayName()
        end if
    end if
End Sub

Dim hsletter
hsletter = 1

Sub HighScoreDisplayName()
	DMDFlush
    Dim i, TempBotStr

    TempBotStr = "  "
    if(hsCurrentDigit> 0) then TempBotStr = TempBotStr & hsEnteredDigits(0)
    if(hsCurrentDigit> 1) then TempBotStr = TempBotStr & hsEnteredDigits(1)
    if(hsCurrentDigit> 2) then TempBotStr = TempBotStr & hsEnteredDigits(2)

    if(hsCurrentDigit <> 3) then
        if(hsLetterFlash <> 0) then
            TempBotStr = TempBotStr & " "
        else
            TempBotStr = TempBotStr & mid(hsValidLetters, hsCurrentLetter, 1)
        end if
    end if

    if(hsCurrentDigit <1) then TempBotStr = TempBotStr & hsEnteredDigits(1)
    if(hsCurrentDigit <2) then TempBotStr = TempBotStr & hsEnteredDigits(2)

    TempBotStr = TempBotStr & "     "
	Select case hsLetter
	Case 0
    DMDId "hsc", "hs.jpg", " Initials:" & mid(TempBotStr, 3, 5), " ",  999999
	hsletter = 27
	Case 1
    DMDId "hsc", "hsa.jpg", " Initials:" & mid(TempBotStr, 3, 5), " ",  999999
	Case 2
    DMDId "hsc", "hsb.jpg", " Initials:" & mid(TempBotStr, 3, 5), " ",  999999
	Case 3
    DMDId "hsc", "hsc.jpg", " Initials:" & mid(TempBotStr, 3, 5), " ",  999999
	Case 4
    DMDId "hsc", "hsd.jpg", " Initials:" & mid(TempBotStr, 3, 5), " ",  999999
	Case 5
    DMDId "hsc", "hse.jpg", " Initials:" & mid(TempBotStr, 3, 5), " ",  999999
	Case 6
    DMDId "hsc", "hsf.jpg", " Initials:" & mid(TempBotStr, 3, 5), " ",  999999
	Case 7
    DMDId "hsc", "hsg.jpg", " Initials:" & mid(TempBotStr, 3, 5), " ",  999999
	Case 8
    DMDId "hsc", "hsh.jpg", " Initials:" & mid(TempBotStr, 3, 5), " ",  999999
	Case 9
    DMDId "hsc", "hsi.jpg", " Initials:" & mid(TempBotStr, 3, 5), " ",  999999
	Case 10
    DMDId "hsc", "hsj.jpg", " Initials:" & mid(TempBotStr, 3, 5), " ",  999999
	Case 11
    DMDId "hsc", "hsk.jpg", " Initials:" & mid(TempBotStr, 3, 5), " ",  999999
	Case 12
    DMDId "hsc", "hsl.jpg", " Initials:" & mid(TempBotStr, 3, 5), " ",  999999
	Case 13
    DMDId "hsc", "hsm.jpg", " Initials:" & mid(TempBotStr, 3, 5), " ",  999999
	Case 14
    DMDId "hsc", "hsn.jpg", " Initials:" & mid(TempBotStr, 3, 5), " ",  999999
	Case 15
    DMDId "hsc", "hso.jpg", " Initials:" & mid(TempBotStr, 3, 5), " ",  999999
	Case 16
    DMDId "hsc", "hsp.jpg", " Initials:" & mid(TempBotStr, 3, 5), " ",  999999
	Case 17
    DMDId "hsc", "hsq.jpg", " Initials:" & mid(TempBotStr, 3, 5), " ",  999999
	Case 18
    DMDId "hsc", "hsr.jpg", " Initials:" & mid(TempBotStr, 3, 5), " ",  999999
	Case 19
    DMDId "hsc", "hss.jpg", " Initials:" & mid(TempBotStr, 3, 5), " ",  999999
	Case 20
    DMDId "hsc", "hst.jpg", " Initials:" & mid(TempBotStr, 3, 5), " ",  999999
	Case 21
    DMDId "hsc", "hsu.jpg", " Initials:" & mid(TempBotStr, 3, 5), " ",  999999
	Case 22
    DMDId "hsc", "hsv.jpg", " Initials:" & mid(TempBotStr, 3, 5), " ",  999999
	Case 23
    DMDId "hsc", "hsw.jpg", " Initials:" & mid(TempBotStr, 3, 5), " ",  999999
	Case 24
    DMDId "hsc", "hsx.jpg", " Initials:" & mid(TempBotStr, 3, 5), " ",  999999
	Case 25
    DMDId "hsc", "hsy.jpg", " Initials:" & mid(TempBotStr, 3, 5), " ",  999999
	Case 26
    DMDId "hsc", "hsz.jpg", " Initials:" & mid(TempBotStr, 3, 5), " ",  999999
	Case 27
    DMDId "hsc", "hs.jpg", " Initials:" & mid(TempBotStr, 3, 5), " ",  999999
	Case 28
    DMDId "hsc", "hsa.jpg", " Initials:" & mid(TempBotStr, 3, 5), " ",  999999
	hsletter = 1
	End Select
    DMDMod "hsc", " Initials:" & mid(TempBotStr, 3, 5), " ", 999999
End Sub

Sub HighScoreCommitName()
    hsbModeActive = False
    PlaySong "m_end"
    hsEnteredName = hsEnteredDigits(0) & hsEnteredDigits(1) & hsEnteredDigits(2)
    if(hsEnteredName = "   ") then
        hsEnteredName = "YOU"
    end if

    HighScoreName(3) = hsEnteredName
    SortHighscore
    DMDFlush
    EndOfBallComplete()
End Sub

Sub SortHighscore
    Dim tmp, tmp2, i, j
    For i = 0 to 3
        For j = 0 to 2
            If HighScore(j) <HighScore(j + 1) Then
                tmp = HighScore(j + 1)
                tmp2 = HighScoreName(j + 1)
                HighScore(j + 1) = HighScore(j)
                HighScoreName(j + 1) = HighScoreName(j)
                HighScore(j) = tmp
                HighScoreName(j) = tmp2
            End If
        Next
    Next
End Sub

'****************************************
' Real Time updates using the GameTimer
'****************************************
'used for all the real time updates

Sub GameTimer_Timer
    RollingUpdate
' add any other real time update subs, like gates or diverters
End Sub

'********************************************************************************************
' Only for VPX 10.2 and higher.
' FlashForMs will blink light or a flasher for TotalPeriod(ms) at rate of BlinkPeriod(ms)
' When TotalPeriod done, light or flasher will be set to FinalState value where
' Final State values are:   0=Off, 1=On, 2=Return to previous State
'********************************************************************************************

Sub FlashForMs(MyLight, TotalPeriod, BlinkPeriod, FinalState) 'thanks gtxjoe for the first version

    If TypeName(MyLight) = "Light" Then

        If FinalState = 2 Then
            FinalState = MyLight.State 'Keep the current light state
        End If
        MyLight.BlinkInterval = BlinkPeriod
        MyLight.Duration 2, TotalPeriod, FinalState
    ElseIf TypeName(MyLight) = "Flasher" Then

        Dim steps

        ' Store all blink information
        steps = Int(TotalPeriod / BlinkPeriod + .5) 'Number of ON/OFF steps to perform
        If FinalState = 2 Then                      'Keep the current flasher state
            FinalState = ABS(MyLight.Visible)
        End If
        MyLight.UserValue = steps * 10 + FinalState 'Store # of blinks, and final state

        ' Start blink timer and create timer subroutine
        MyLight.TimerInterval = BlinkPeriod
        MyLight.TimerEnabled = 0
        MyLight.TimerEnabled = 1
        ExecuteGlobal "Sub " & MyLight.Name & "_Timer:" & "Dim tmp, steps, fstate:tmp=me.UserValue:fstate = tmp MOD 10:steps= tmp\10 -1:Me.Visible = steps MOD 2:me.UserValue = steps *10 + fstate:If Steps = 0 then Me.Visible = fstate:Me.TimerEnabled=0:End if:End Sub"
    End If
End Sub

'******************************************
' Change light color - simulate color leds
' changes the light color and state
' 10 colors: red, orange, amber, yellow...
'******************************************
' in this table this colors are use to keep track of the progress during the acts and battles

'colors
Dim red, orange, amber, yellow, darkgreen, green, blue, darkblue, purple, white

red = 10
orange = 9
amber = 8
yellow = 7
darkgreen = 6
green = 5
blue = 4
darkblue = 3
purple = 2
white = 1

Sub SetLightColor(n, col, stat)
    Select Case col
        Case red
            n.color = RGB(18, 0, 0)
            n.colorfull = RGB(255, 0, 0)
        Case orange
            n.color = RGB(18, 3, 0)
            n.colorfull = RGB(255, 64, 0)
        Case amber
            n.color = RGB(193, 49, 0)
            n.colorfull = RGB(255, 153, 0)
        Case yellow
            n.color = RGB(18, 18, 0)
            n.colorfull = RGB(255, 255, 0)
        Case darkgreen
            n.color = RGB(0, 8, 0)
            n.colorfull = RGB(0, 64, 0)
        Case green
            n.color = RGB(0, 18, 0)
            n.colorfull = RGB(0, 255, 0)
        Case blue
            n.color = RGB(0, 18, 18)
            n.colorfull = RGB(0, 255, 255)
        Case darkblue
            n.color = RGB(0, 8, 8)
            n.colorfull = RGB(0, 64, 64)
        Case purple
            n.color = RGB(128, 0, 128)
            n.colorfull = RGB(255, 0, 255)
        Case white
            n.color = RGB(255, 252, 224)
            n.colorfull = RGB(193, 91, 0)
        Case white
            n.color = RGB(255, 252, 224)
            n.colorfull = RGB(193, 91, 0)
    End Select
    If stat <> -1 Then
        n.State = 0
        n.State = stat
    End If
End Sub

Sub ResetAllLightsColor ' Called at a new game
    'shoot again
    SetLightColor LightShootAgain, red, -1
    ' lanes
	SetLightColor ll1, blue, -1
	SetLightColor ll2, blue, -1
	SetLightColor ll3, blue, -1
	SetLightColor ll4, blue, -1
	SetLightColor ll5, blue, -1
	' Run Lights
	SetLightColor lr1, red, -1
	SetLightColor lr2, red, -1
	SetLightColor lr3, red, -1
	' Mode Lights
	SetLightColor lm1, yellow, -1
	SetLightColor lm2, red, -1
	SetLightColor lm3, blue, -1
	SetLightColor lm4, red, -1
	' escape lights row 1
	SetLightColor le2, blue, -1
	SetLightColor le4, blue, -1
	SetLightColor le6, blue, -1
	SetLightColor le11, blue, -1
	SetLightColor le13, blue, -1
	SetLightColor le15, blue, -1
	' escape lights row 2
	SetLightColor le1, purple, -1
	SetLightColor le3, purple, -1
	SetLightColor le5, purple, -1
	SetLightColor le10, purple, -1
	SetLightColor le12, purple, -1
	SetLightColor le14, purple, -1
	' escape Go Lights
	SetLightColor le7, yellow, -1
	SetLightColor le8, orange, -1
	SetLightColor le9, red, -1
	' barb Lights
	SetLightColor lb1, yellow, -1
	SetLightColor lb2, yellow, -1
	SetLightColor lb3, yellow, -1
	SetLightColor lb4, yellow, -1
	' lock Lights
	SetLightColor llo2, yellow, -1
	SetLightColor llo4, red, -1
	SetLightColor lro3, yellow, -1
	SetLightColor lc2, blue, -1
	'Extra Ball 
	SetLightColor llo5, orange, -1
	SetLightColor lro1, orange, -1
	' Orbit & Ramp Lights
	SetLightColor llo1, red, -1
	SetLightColor llr1, red, -1
	SetLightColor llo3, red, -1
	SetLightColor lc1, red, -1
	SetLightColor lro2, red, -1
	SetLightColor lrr1, red, -1
	SetLightColor lro4, red, -1
End Sub

Sub UpdateBonusColors
End Sub

'*************************
' Rainbow Changing Lights
'*************************

Dim RGBStep, RGBFactor, rRed, rGreen, rBlue, RainbowLights

Sub StartRainbow(n)
    set RainbowLights = n
    RGBStep = 0
    RGBFactor = 5
    rRed = 255
    rGreen = 0
    rBlue = 0
    RainbowTimer.Enabled = 1
End Sub

Sub StopRainbow()
    Dim obj
    RainbowTimer.Enabled = 0
    RainbowTimer.Enabled = 0
        For each obj in RainbowLights
            SetLightColor obj, "white", 0
        Next
End Sub

Sub RainbowTimer_Timer 'rainbow led light color changing
    Dim obj
    Select Case RGBStep
        Case 0 'Green
            rGreen = rGreen + RGBFactor
            If rGreen > 255 then
                rGreen = 255
                RGBStep = 1
            End If
        Case 1 'Red
            rRed = rRed - RGBFactor
            If rRed < 0 then
                rRed = 0
                RGBStep = 2
            End If
        Case 2 'Blue
            rBlue = rBlue + RGBFactor
            If rBlue > 255 then
                rBlue = 255
                RGBStep = 3
            End If
        Case 3 'Green
            rGreen = rGreen - RGBFactor
            If rGreen < 0 then
                rGreen = 0
                RGBStep = 4
            End If
        Case 4 'Red
            rRed = rRed + RGBFactor
            If rRed > 255 then
                rRed = 255
                RGBStep = 5
            End If
        Case 5 'Blue
            rBlue = rBlue - RGBFactor
            If rBlue < 0 then
                rBlue = 0
                RGBStep = 0
            End If
    End Select
        For each obj in RainbowLights
            obj.color = RGB(rRed \ 10, rGreen \ 10, rBlue \ 10)
            obj.colorfull = RGB(rRed, rGreen, rBlue)
        Next
End Sub

'***********************************************************************************
'         	    JPS DMD - very simple DMD routines using UltraDMD
'***********************************************************************************

Dim UltraDMD

' DMD using UltraDMD calls

Sub DMD(background, toptext, bottomtext, duration)
    UltraDMD.DisplayScene00 background, toptext, 15, bottomtext, 15, 14, duration, 14
    UltraDMDTimer.Enabled = 1 'to show the score after the animation/message
End Sub

Sub DMDScore
	If inconfig = True Then
    UltraDMD.SetScoreboardBackgroundImage "black.png", 15, 7
    UltraDMD.DisplayScoreboard PlayersPlayingGame, CurrentPlayer, Score(1), Score(2), Score(3), Score(4), "" & CurrentPlayer, "Hold Flippers to Exit"
	Else
    UltraDMD.SetScoreboardBackgroundImage "scoreboard-background.jpg", 15, 7
    UltraDMD.DisplayScoreboard PlayersPlayingGame, CurrentPlayer, Score(1), Score(2), Score(3), Score(4), "Player " & CurrentPlayer, "Ball " & Balls
	End If
End Sub

Sub DMDScoreNow
    DMDFlush
    DMDScore
End Sub

Sub DMDFLush
    UltraDMDTimer.Enabled = 0
    UltraDMD.CancelRendering
End Sub

Sub DMDScrollCredits(background, text, duration)
    UltraDMD.ScrollingCredits background, text, 15, 14, duration, 14
End Sub

Sub DMDId(id, background, toptext, bottomtext, duration)
    UltraDMD.DisplayScene00ExwithID id, False, background, toptext, 15, 0, bottomtext, 15, 0, 14, duration, 14
End Sub

Sub DMDMod(id, toptext, bottomtext, duration)
    UltraDMD.ModifyScene00Ex id, toptext, bottomtext, duration
End Sub

Sub UltraDMDTimer_Timer 'used for the attrackmode and the instant info.
    If bInstantInfo Then
        InstantInfo
    ElseIf bAttractMode Then
        ShowTableInfo
    ElseIf NOT UltraDMD.IsRendering Then
        DMDScoreNow
    ElseIf bromconfig Then
        romconfig
    End If
End Sub

Sub DMD_Init
    Set UltraDMD = CreateObject("UltraDMD.DMDObject")
    If UltraDMD is Nothing Then
        MsgBox "No UltraDMD found.  This table will NOT run without it."
        Exit Sub
    End If

    UltraDMD.Init
    If Not UltraDMD.GetMajorVersion = 1 Then
        MsgBox "Incompatible Version of UltraDMD found."
        Exit Sub
    End If

    If UltraDMD.GetMinorVersion <1 Then
        MsgBox "Incompatible Version of UltraDMD found. Please update to version 1.1 or newer."
        Exit Sub
    End If

    Dim fso:Set fso = CreateObject("Scripting.FileSystemObject")
    Dim curDir:curDir = fso.GetAbsolutePathName(".")

    Dim DirName
    DirName = curDir& "\" &TableName& ".UltraDMD"

    If Not fso.FolderExists(DirName) Then _
            Msgbox "UltraDMD userfiles directory '" & DirName & "' does not exist." & CHR(13) & "No graphic images will be displayed on the DMD"
    UltraDMD.SetProjectFolder DirName

    ' wait for the animation to end
    While UltraDMD.IsRendering = True
    WEnd

End Sub

' ********************************
'   Table info & Attract Mode
' ********************************


Sub ShowTableInfo
	'dmdintroloop
End Sub

Dim introposition
introposition = 0

Sub dmdattract_timer
	If NOT UltraDMD.IsRendering Then
	dmdintroloop
	End If
End Sub

Sub dmdintroloop
	DMDFlush
	introposition = introposition + 1
	Select Case introposition
	Case 1
	Dim i
		If Score(1) Then
		DMD "eggo1-17.wmv", "", "",  18000
			DMD "black.png", "PLAYER 1", Score(1), 3000
		Else
		introposition = 4
		dmdintroloop
		End If
	Case 2
		If Score(2) Then
			DMD "black.png", "PLAYER 2", Score(2), 3000
		Else
		introposition = 4
		dmdintroloop
		End If
	Case 3
		If Score(3) Then
			DMD "black.png", "PLAYER 3", Score(3), 3000
		Else
		introposition = 4
		dmdintroloop
		End If
	Case 4
		If Score(4) Then
			DMD "black.png", "PLAYER 4", Score(4), 3000
		Else
		introposition = 4
		dmdintroloop
		End If
	Case 5
		'coins or freeplay
		If gamemodecurrent = "True" Then
			DMD "black.png", " ", "FREE PLAY",  2000
		Else
			If Credits > 0 Then
				DMD "black.png", "CREDITS " &credits, "PRESS START",  2000
				DOF 140, DOFOn
			Else
				DMD "black.png", "CREDITS " &credits, "INSERT COIN",  2000
				DOF 140, DOFOff
			End If
		End If
	Case 6
		DMD "introconfig.png", "", "",  5000
	Case 7
		DMD "intro-49.wmv", "", "",  48000
	Case 8
    DMD "black.png", "Grand Champion", " " & HighScoreName(0) & " " & FormatNumber(HighScore(0),0,,, -1), 3000
	Case 9
    DMD "black.png", "HIGHSCORE 1", "" & HighScoreName(1) & " " & FormatNumber(HighScore(1),0,,, -1), 3000
	Case 10
    DMD "black.png", "HIGHSCORE 2", "" & HighScoreName(2) & " " & FormatNumber(HighScore(2),0,,, -1), 3000
	Case 11
    DMD "black.png", "HIGHSCORE 3", "" & HighScoreName(3) & " " & FormatNumber(HighScore(3),0,,, -1), 3000
	introposition = 0
End Select
End Sub

Sub StartAttractMode()
    bAttractMode = True
    UltraDMDTimer.Enabled = 1
    StartLightSeq
    'ShowTableInfo
	'dmdintroloop
    StartRainbow aLights
	inconfig = false
	dmdattract.Enabled = 1
End Sub

Sub StopAttractMode()
    bAttractMode = False
    DMDScoreNow
    LightSeqAttract.StopPlay
    LightSeqFlasher.StopPlay
    StopRainbow
    ResetAllLightsColor
	dmdattract.Enabled = 0
'StopSong
End Sub

Sub StartLightSeq()
    'lights sequences
	LightSeqaxmas.UpdateInterval = 150
	LightSeqaxmas.Play SeqRandom, 10, , 50000
    LightSeqFlasher.UpdateInterval = 150
    LightSeqFlasher.Play SeqRandom, 10, , 50000
    LightSeqAttract.UpdateInterval = 25
    LightSeqAttract.Play SeqBlinking, , 5, 150
    LightSeqAttract.Play SeqRandom, 40, , 4000
    LightSeqAttract.Play SeqAllOff
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqUpOn, 50, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqDownOn, 25, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqCircleOutOn, 15, 2
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqUpOn, 25, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqDownOn, 25, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqUpOn, 25, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqDownOn, 25, 1
    LightSeqAttract.UpdateInterval = 10
    LightSeqAttract.Play SeqCircleOutOn, 15, 3
    LightSeqAttract.UpdateInterval = 5
    LightSeqAttract.Play SeqRightOn, 50, 1
    LightSeqAttract.UpdateInterval = 5
    LightSeqAttract.Play SeqLeftOn, 50, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqRightOn, 50, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqLeftOn, 50, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqRightOn, 40, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqLeftOn, 40, 1
    LightSeqAttract.UpdateInterval = 10
    LightSeqAttract.Play SeqRightOn, 30, 1
    LightSeqAttract.UpdateInterval = 10
    LightSeqAttract.Play SeqLeftOn, 30, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqRightOn, 25, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqLeftOn, 25, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqRightOn, 15, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqLeftOn, 15, 1
    LightSeqAttract.UpdateInterval = 10
    LightSeqAttract.Play SeqCircleOutOn, 15, 3
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqLeftOn, 25, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqRightOn, 25, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqLeftOn, 25, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqUpOn, 25, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqDownOn, 25, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqUpOn, 25, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqDownOn, 25, 1
    LightSeqAttract.UpdateInterval = 5
    LightSeqAttract.Play SeqStripe1VertOn, 50, 2
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqCircleOutOn, 15, 2
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqStripe1VertOn, 50, 3
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqLeftOn, 25, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqRightOn, 25, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqLeftOn, 25, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqUpOn, 25, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqDownOn, 25, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqCircleOutOn, 15, 2
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqStripe2VertOn, 50, 3
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqLeftOn, 25, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqRightOn, 25, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqLeftOn, 25, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqUpOn, 25, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqDownOn, 25, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqUpOn, 25, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqDownOn, 25, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqStripe1VertOn, 25, 3
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqStripe2VertOn, 25, 3
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqUpOn, 15, 1
    LightSeqAttract.UpdateInterval = 8
	LightSeqAttract.Play SeqDownOn, 15, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqUpOn, 15, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqDownOn, 15, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqUpOn, 15, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqDownOn, 15, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqRightOn, 15, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqLeftOn, 15, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqRightOn, 15, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqLeftOn, 15, 1
End Sub

Sub LightSeqAttract_PlayDone()
    StartLightSeq()
End Sub

Sub LightSeqTilt_PlayDone()
    LightSeqTilt.Play SeqAllOff
End Sub

Sub LightSeqSkillshot_PlayDone()
    LightSeqSkillshot.Play SeqAllOff
End Sub




'***********************************************************************
' *********************************************************************
'                     Table Specific Script Starts Here
' *********************************************************************
'***********************************************************************
' droptargets, animations, etc
Sub VPObjects_Init

End Sub

' tables variables and Mode init
Dim LaneBonus
Dim TargetBonus
Dim RampBonus
Dim OrbitBonus
Dim spinvalue
Dim barbMultiball
Dim barbHits(10)
Dim LookForBarb
Dim RunMultiball
Dim RunAway
Dim RunHits(9)
Dim BallsInRunLock(4)
Dim LRHits(10)
Dim RRHits(7)
Dim WillMultiball
Dim WillHits(3)
Dim WillSuper
Dim WillSuperReady
Dim DemoMultiball
Dim DemoHits
Dim MonsterFinalBlow
Dim BarbJackpots
Dim finalflips

Sub Game_Init() 'called at the start of a new game
    Dim i
	inconfig = false
	lrflashtime.Enabled = False
    bExtraBallWonThisBall = False
   ' TurnOffPlayfieldLights()
    'Play some Music
	CurrentSong
	'Init Variables
	For i = 0 to 4
        SkillshotValue(i) = 1000000 ' increases by 1000000 each time it is collected
    Next
    'Init Delays/Timers
    PlayQuote.Enabled = 1

    'MainMode Init()
    'For i = 0 to 10
    '    Mode(i) = 0
    'Next
	MagnetB.MagnetON = False
	lowerflippersoff = True
	CloseGates

	'Init lights
	'Barb Multiball resets 
	barbMultiball = false
For i = 0 to 10
	barbHits(i) = 0
Next
	LookForBarb = False
	' Run Multiball resets
	RunMultiball = False
	RunAway = False
For i = 0 to 9
	RunHits(i) = 0
Next
	finalflips = False
For i = 0 to 4
	BallsInRunLock(i) = 0
Next
	' Will Resets  
	WillMultiball = False
For i = 0 to 3
	WillHits(i) = 0
Next
	WillSuperReady = False
	'Monster Resets  
	DemoMultiball = False
	DemoHits = 0
	MonsterFinalBlow = False
	BarbJackpots = False
	' Reamp Resets  
For i = 0 to 10
	LRHits(i) = 0
Next
For i = 0 to 7
	RRHits(i) = 0
Next
For i = 0 to 100
	bumps(i) = 0
Next
For i = 0 to 25
	SteveHit(i) = 0
Next
For i = 0 to 25
	NancyHit(i) = 0
Next
For i = 0 to 100
	orbithit(i) = 0
Next
End Sub

Sub StopEndOfBallMode()              'this sub is called after the last ball is drained
    ResetSkillShotTimer_Timer
End Sub

Sub ResetNewBallVariables()          'reset variables for a new ball or player

End Sub

Sub ResetNewBallLights() 'turn on or off the needed lights before a new ball is released

End Sub

Sub TurnOffPlayfieldLights()
    Dim a
    For each a in aLights
        a.State = 0
    Next
End Sub

Sub UpdateSkillShot() 'Updates the skillshot light
    LightSeqSkillshot.Play SeqAllOff
    lr1.State = 2
    lr2.State = 2
    lr3.State = 2
'DMDFlush
End Sub

Sub SkillshotOff_Hit 'trigger to stop the skillshot due to a weak plunger shot
    If bSkillShotReady Then
        ResetSkillShotTimer_Timer
    End If
End Sub

Sub ResetSkillShotTimer_Timer 'timer to reset the skillshot lights & variables
    ResetSkillShotTimer.Enabled = 0
    bSkillShotReady = False
    LightSeqSkillshot.StopPlay
    If lr1.State = 2 Then lr1.State = 0
    If lr2.State = 2 Then lr2.State = 0
    If lr3.State = 2 Then lr3.State = 0
'DMDScoreNow
End Sub


'*******************************************
'******************************
' *********************************************************************
'                        Table Object Hit Events
'
' Any target hit Sub will follow this:
' - play a sound
' - do some physical movement
' - add a score, bonus
' - check some variables/Mode this trigger is a member of
' - set the "LastSwitchHit" variable in case it is needed later
' *********************************************************************

' Slingshots has been hit

Dim LStep, RStep

Sub LeftSlingShot_Slingshot
    If Tilted Then Exit Sub
	startB2S(1)
	LightEffect 7
    PlaySound SoundFXDOF("fx_slingshot", 103, DOFPulse, DOFContactors), 0, 1, -0.05, 0.05
	PlaySound "whomp"
    DOF 104, DOFPulse
    LeftSling4.Visible = 1
    Lemk.RotX = 26
    LStep = 0
    LeftSlingShot.TimerEnabled = True
    ' add some points
    AddScore 110
    ' add some effect to the table?
    'FlashForMs l20, 1000, 50, 0:FlashForMs l20f, 1000, 50, 0
    'FlashForMs l21, 1000, 50, 0:FlashForMs l21f, 1000, 50, 0
    ' remember last trigger hit by the ball
    LastSwitchHit = "LeftSlingShot"
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
    If Tilted Then Exit Sub
	startB2S(3)
	LightEffect 6
    PlaySound SoundFXDOF("fx_slingshot", 105, DOFPulse, DOFContactors), 0, 1, 0.05, 0.05
	PlaySound "whomp"
    DOF 106, DOFPulse
    RightSling4.Visible = 1
    Lemk1.RotX = 26
    RStep = 0
    RightSlingShot.TimerEnabled = True
    ' add some points
    AddScore 110
    ' add some effect to the table?
    'FlashForMs l22, 1000, 50, 0:FlashForMs l22f, 1000, 50, 0
    'FlashForMs l23, 1000, 50, 0:FlashForMs l23f, 1000, 50, 0
    ' remember last trigger hit by the ball
    LastSwitchHit = "RightSlingShot"
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 1:RightSLing4.Visible = 0:RightSLing3.Visible = 1:Lemk1.RotX = 14
        Case 2:RightSLing3.Visible = 0:RightSLing2.Visible = 1:Lemk1.RotX = 2
        Case 3:RightSLing2.Visible = 0:Lemk1.RotX = -10:RightSlingShot.TimerEnabled = 0
    End Select

    RStep = RStep + 1
End Sub

'*********
' Gates
'*********
' lets get these gates all closed up.
	Gate6.Open = False
	Gate4.Open = False
	Gate8.Open = False
	Gate7.Open = False

Sub CloseGates 
	Gate6.Open = False
	Gate4.Open = False
	Gate8.Open = False
	Gate7.Open = False
End Sub

Sub OpenGates 
	Gate6.Open = True
	Gate4.Open = True
	Gate8.Open = True
	Gate7.Open = True
End Sub

'***************************
'block the upside down during multiballs
'***************************

Sub udblock_hit
Gate7.Open = True
Gate9.Open = True
If barbMultiball = True Then
Gate9.Open = False
End If
If RunMultiball = True Then
Gate9.Open = False
End If
End Sub

'*********
' Bumpers
'*********

Dim bumps(100)
Sub Bumper1_Hit
	WaffleShake()
	LightEffect 5
	startB2S(2)
    If NOT Tilted Then
	bumps(CurrentPlayer) = bumps(CurrentPlayer) + 1
		PlaySound SoundFXDOF("fx_bumper", 107, DOFPulse, DOFContactors), 0, 1, pan(ActiveBall)
	    DOF 110, DOFPulse
		PlaySound "wave"
        ' add some points
        AddScore 1000
        LastSwitchHit = "Bumper1"
	If bMultiBallMode = False Then
	DMD "eggospin.wmv", "EGGOS", bumps(CurrentPlayer),  300
	End If
    End If
End Sub

Sub Bumper2_Hit
	WaffleShake2()
	LightEffect 5
	startB2S(4)
    If NOT Tilted Then
	bumps(CurrentPlayer) = bumps(CurrentPlayer) + 1
	    PlaySound SoundFXDOF("fx_bumper", 109, DOFPulse, DOFContactors), 0, 1, pan(ActiveBall)
		DOF 111, DOFPulse
		PlaySound "wave"
        ' add some points
        AddScore 1000
        LastSwitchHit = "Bumper2"
	If bMultiBallMode = False Then
	DMD "eggospin.wmv", "EGGOS", bumps(CurrentPlayer),  300
	End If
    End If
End Sub

Sub Bumper3_Hit
	WaffleShake3()
	LightEffect 5
	startB2S(5)
    If NOT Tilted Then
	bumps(CurrentPlayer) = bumps(CurrentPlayer) + 1
        PlaySound SoundFXDOF("fx_bumper", 108, DOFPulse, DOFContactors), 0, 1, pan(ActiveBall)
		DOF 111, DOFPulse
		PlaySound "wave"
        ' add some points
        AddScore 1000
        LastSwitchHit = "Bumper1"
	If bMultiBallMode = False Then
	DMD "eggospin.wmv", "EGGOS", bumps(CurrentPlayer),  300
	End If
    End If
End Sub

Sub Bumper4_Hit
	WaffleShake4()
	LightEffect 5
	startB2S(6)
    If NOT Tilted Then
	bumps(CurrentPlayer) = bumps(CurrentPlayer) + 1
        PlaySound SoundFXDOF("fx_bumper", 137, DOFPulse, DOFContactors), 0, 1, pan(ActiveBall)
		DOF 113, DOFPulse
		PlaySound "wave"
        ' add some points
        AddScore 1000
        LastSwitchHit = "Bumper1"
	If bMultiBallMode = False Then
	DMD "eggospin.wmv", "EGGOS", bumps(CurrentPlayer),  300
	End If
    End If
End Sub

dicetime.enabled = 0
Sub dicespin
	dicetime.enabled = 1
	DOF 116, DOFOn
	startB2S(4)
End Sub
Sub dicetime_Timer 
	dice.ObjRotZ = (dice.ObjRotZ + 1) Mod 360
	If dice.ObjRotZ = 358 Then
		dicetime.enabled = 0
		DOF 116, DOFOff
		dice.ObjRotZ = (dice.ObjRotZ + 1) Mod 360
	End If
End Sub

Sub Gate5_Hit
	dicespin
	LightEffect 11
    DOF 141, DOFOff
End Sub

Sub BumperRewards
		Select Case bumps(CurrentPlayer)
			Case 1
				DMD "black.png", "Collect EGGOS", "Rewards at 25,50,75,100",  1000
            Case 25
				DMD "eggo2-4.wmv", "", "", 5000
				AddScore 1000000
                DMD "black.png", "MMMMMM EGGOS", "1000000 Awarded",  1000
            Case 50
				DMD "eggo2-4.wmv", "", "", 5000
				AddScore 2000000
                DMD "black.png", "Super EGGOS", "2000000 Awarded",  1000
            Case 75
				DMD "eggo2-4.wmv", "", "", 5000
				AddScore 4000000
                DMD "black.png", "Double Super EGGOS", "4000000 Awarded",  1000
            Case 99
				DMD "eggo2-4.wmv", "", "", 5000
				AddScore 6000000
                DMD "black.png", "Ultra EGGOS", "6000000 Awarded",  1000

        End Select
End Sub
'**********************
' Flipper Lanes: The Kids
'**********************

Sub lane1_Hit
    PlaySound "fx_sensor", 0, 1, pan(ActiveBall)
	If bMultiBallMode = False Then
		PlaySound "lane"
		DOF 144, DOFPulse
	End If
    If Tilted Then Exit Sub
    LaneBonus = LaneBonus + 1
    ll1.State = 1
    AddScore 50050
    ' Do some sound or light effect

    LastSwitchHit = "lane1"
    ' do some check
    CheckKIDSLane
	startB2S(1)
End Sub

Sub lane2_Hit
    PlaySound "fx_sensor", 0, 1, pan(ActiveBall)
	If bMultiBallMode = False Then
		PlaySound "lane"
		DOF 145, DOFPulse
	End If
    If Tilted Then Exit Sub
    LaneBonus = LaneBonus + 1
    ll2.State = 1
    AddScore 10010
    ' Do some sound or light effect

    LastSwitchHit = "lane2"
    ' do some check
    CheckKIDSLane
	startB2S(2)
End Sub

Sub lane3_Hit
    PlaySound "fx_sensor", 0, 1, pan(ActiveBall)
	If bMultiBallMode = False Then
		PlaySound "lane"
		DOF 146, DOFPulse
	End If
    If Tilted Then Exit Sub
    LaneBonus = LaneBonus + 1
    ll3.State = 1
    AddScore 10010
    ' Do some sound or light effect

    LastSwitchHit = "lane3"
    ' do some check
    CheckKIDSLane
	startB2S(3)
End Sub

Sub lane4_Hit
    PlaySound "fx_sensor", 0, 1, pan(ActiveBall)
	If bMultiBallMode = False Then
		PlaySound "lane"
		DOF 147, DOFPulse
	End If
    If Tilted Then Exit Sub
    LaneBonus = LaneBonus + 1
    ll4.State = 1
    AddScore 10010
    ' Do some sound or light effect

    LastSwitchHit = "lane4"
    ' do some check
    CheckKIDSLane
	startB2S(4)
End Sub

Sub lane5_Hit
    PlaySound "fx_sensor", 0, 1, pan(ActiveBall)
	If bMultiBallMode = False Then
		PlaySound "lane"
		DOF 148, DOFPulse
	End If
    If Tilted Then Exit Sub
    LaneBonus = LaneBonus + 1
    ll5.State = 1
    AddScore 50050
    ' Do some sound or light effect

    LastSwitchHit = "lane5"
    ' do some check
    CheckKIDSLane
	startB2S(4)
End Sub

Sub CheckKIDSLane() 'use the lane lights
	If bMultiBallMode = True Then Exit Sub
    If ll1.State + ll2.State + ll3.State + ll4.State + ll5.State = 5 Then
        'Activate Ball Saver
		EnableBallSaver 15
        ResetKidsLights
		DMD "black.png", "Ball Save","Activated",  500

    End If
End Sub

Sub ResetKidsLights
	ll1.State = 0
	ll2.State = 0
	ll3.State = 0
	ll4.State = 0
	ll5.State = 0
End Sub

'*****************
'  BARB Targets
'*****************

Sub barb1_Hit
	LightEffect 10
	startB2S(6)
	DOF 133, DOFPulse
    PlaySound SoundFX("fx_target",DOFTargets), 0, 1, pan(ActiveBall)
	PlaySound "runtarget"
    If Tilted Then Exit Sub
    lb1.State = 1
    AddScore 25010
    ' Do some sound or light effect
    FlashForMs f4, 1000, 50, 0
    FlashForMs f10, 1000, 50, 0
    LastSwitchHit = "barb1"
    ' do some check
    CheckBARBTargets
End Sub

Sub barb2_Hit
	LightEffect 10
	startB2S(6)
	DOF 133, DOFPulse
    PlaySound SoundFX("fx_target",DOFTargets), 0, 1, pan(ActiveBall)
	PlaySound "runtarget"
    If Tilted Then Exit Sub
    lb2.State = 1
    AddScore 25010
    ' Do some sound or light effect
    FlashForMs f4, 1000, 50, 0
    FlashForMs f10, 1000, 50, 0
    LastSwitchHit = "barb2"
    ' do some check
    CheckBARBTargets
End Sub

Sub barb3_Hit
	LightEffect 10
	startB2S(6)
	DOF 133, DOFPulse
    PlaySound SoundFX("fx_target",DOFTargets), 0, 1, pan(ActiveBall)
	PlaySound "runtarget"
    If Tilted Then Exit Sub

    lb3.State = 1
    AddScore 25010
    ' Do some sound or light effect
    FlashForMs f4, 1000, 50, 0
    FlashForMs f10, 1000, 50, 0
    LastSwitchHit = "barb3"
    ' do some check
    CheckBARBTargets
End Sub

Sub barb4_Hit
	LightEffect 10
	startB2S(6)
	DOF 133, DOFPulse
    PlaySound SoundFX("fx_target",DOFTargets), 0, 1, pan(ActiveBall)
	PlaySound "runtarget"
    If Tilted Then Exit Sub

    lb4.State = 1
    AddScore 25010
    ' Do some sound or light effect
    LastSwitchHit = "barb4"
    ' do some check
    CheckBARBTargets
End Sub

Sub CheckBARBTargets
If bMultiBallMode = True Then Exit Sub
If lb1.State + lb2.State + lb3.State + lb4.State = 4 Then
	Gate4.Open = True
	MagnetB.MagnetON = True ' Magnet On
	PlaySound "barblock"
	DMD "barb-r4-3.wmv", "", "", 4000
	DMD "black.png", "Barb Lock","is Lit",  500
	llo1.State = 2
	lro2.State = 2
	lro3.State = 2
	llo2.State = 2
End If
End Sub

Sub BallLockBarb_Hit
	startB2S(6)
    Dim waittime
    waittime = 100
	MagnetB.MagnetON = False ' Magnet no longer needed
	If LookForBarb = True Then
	BarbSuper
	Else
	BallsInLock(CurrentPlayer) = BallsInLock(CurrentPlayer) + 1
        Select Case BallsInLock(CurrentPlayer)
            Case 1
				ResetBARBLights
				DMD "barb-r1-4.wmv", "", "", 5000
                DMD "black.png", "Barb","Ball 1 Locked",  500
                waittime = 1000
			    vpmtimer.addtimer waittime, "BallLockBarbExit '"
            Case 2
				ResetBARBLights
				DMDFlush
                waittime = 18000
				DMD "barb-start-17.wmv", "", "", 18000
				DMD "black.png", "Where's Barb", "Multiball", 2000
				vpmtimer.addtimer waittime, "StartBarb'"
				vpmtimer.addtimer waittime, "BallLockBarbExit '"
        End Select
	End If
End Sub

Sub BarbUp
		If barb1.IsDropped = 1 or barb2.IsDropped = 1 or barb3.IsDropped = 1 or barb4.IsDropped = 1 Then DOF 129, DOFPulse
	barb1.IsDropped = 0
	barb2.IsDropped = 0
	barb3.IsDropped = 0
	barb4.IsDropped = 0
	CheckBARBTargets
End Sub

Sub ResetBARBLights
		If barb1.IsDropped = 1 or barb2.IsDropped = 1 or barb3.IsDropped = 1 or barb4.IsDropped = 1 Then DOF 129, DOFPulse
	barb1.IsDropped = 0
	barb2.IsDropped = 0
	barb3.IsDropped = 0
	barb4.IsDropped = 0
lb1.State = 0
lb2.State = 0
lb3.State = 0
lb4.State = 0
llo1.State = 0
lro2.State = 0
lro3.State = 0
llo2.State = 0
End Sub

'*****************
'  RUN Targets
'*****************

Sub run1_Hit
	LightEffect 9
	startB2S(5)
	DOF 134, DOFPulse
	TruckShake()
	If bSkillShotReady Then
	AwardSkillshot
	End If
    PlaySound SoundFX("fx_target",DOFTargets), 0, 1, pan(ActiveBall)
	PlaySound "runtarget"
    If Tilted Then Exit Sub

    lr1.State = 1
    AddScore 25010
    ' Do some sound or light effect
    FlashForMs f1, 1000, 50, 0
    LastSwitchHit = "run1"
    ' do some check
    CheckRunTargets
End Sub

Sub run2_Hit
	LightEffect 9
	startB2S(5)
	DOF 134, DOFPulse
	TruckShake()
	If bSkillShotReady Then
	AwardSkillshot
	End If
    PlaySound SoundFX("fx_target",DOFTargets), 0, 1, pan(ActiveBall)
	PlaySound "runtarget"
    If Tilted Then Exit Sub

    lr2.State = 1
    AddScore 25010
    ' Do some sound or light effect
    FlashForMs f1, 1000, 50, 0
    LastSwitchHit = "run2"
    ' do some check
    CheckRunTargets
End Sub

Sub run3_Hit
	LightEffect 9
	startB2S(5)
	DOF 134, DOFPulse
	TruckShake()
	If bSkillShotReady Then
	AwardSkillshot
	End If
    PlaySound SoundFX("fx_target",DOFTargets), 0, 1, pan(ActiveBall)
	PlaySound "runtarget"
    If Tilted Then Exit Sub

    lr3.State = 1
    AddScore 25010
    ' Do some sound or light effect
    FlashForMs f1, 1000, 50, 0
    LastSwitchHit = "run3"
    ' do some check
    CheckRunTargets
End Sub


Sub CheckRunTargets
If bMultiBallMode = True Then Exit Sub
If lr1.State + lr2.State + lr3.State = 3 Then
	Gate6.Open = True
	PlaySound "ping"
	DMD "lock-3.wmv", "", "", 4000
	DMD "black.png", "Run Lock","is Lit",  500
	llo4.State = 2
	llo3.State = 2
End If
End Sub

Sub BallLockRun_Hit
    Dim waittime
    waittime = 100
	If RunAway = True Then
	RunSuper
	Else
	BallsInRunLock(CurrentPlayer) = BallsInRunLock(CurrentPlayer) + 1
        Select Case BallsInRunLock(CurrentPlayer)
            Case 1
				ResetRunLights
				DMD "r5-3.wmv", "", "", 4000
                DMD "black.png", "Run","Ball 1 Locked",  500
                waittime = 1000
				vpmtimer.addtimer waittime, "BallLockRunExit '"
            Case 2
				ResetRunLights
                waittime = 22000
				DMDFlush
				DMD "run-start-21.wmv", "","",  22000
				DMD "black.png", "RUN RUN RUN", "Multiball", 1000
				vpmtimer.addtimer waittime, "StartRun'"
				vpmtimer.addtimer waittime, "BallLockRunExit '"
				
        End Select
	End If
    'vpmtimer.addtimer waittime, "BallLockRunExit '"
End Sub

Sub RunUp
		If run1.IsDropped = 1 or run2.IsDropped = 1 or run3.IsDropped = 1 Then DOF 130, DOFPulse
	run1.IsDropped = 0
	run2.IsDropped = 0
	run3.IsDropped = 0
	CheckRunTargets
End Sub

Sub ResetRunLights
	If run1.IsDropped = 1 or run2.IsDropped = 1 or run3.IsDropped = 1 Then DOF 130, DOFPulse
	run1.IsDropped = 0
	run2.IsDropped = 0
	run3.IsDropped = 0
lr1.State = 0
lr2.State = 0
lr3.State = 0
llo4.State = 0
llo3.State = 0
End Sub

'***************************
'Hawkins Electric Truck Animation / shake
'***************************

Dim TruckPos

Sub TruckShake()
	startB2S(2)
    TruckPos = 3
	DOF 128, DOFOn
    TruckShakeTimer.Enabled = True
End Sub

Sub TruckShakeTimer_Timer()
    Hawkins.RotY = TruckPos
    If TruckPos <= 0.1 AND TruckPos >= -0.1 Then Me.Enabled = False:DOF 128, DOFOff:Exit Sub
    If TruckPos < 0 Then
        TruckPos = ABS(TruckPos)- 0.1
    Else
        TruckPos = - TruckPos + 0.1
    End If
End Sub

'***************************
'Waffle Animation / shake
'***************************

Dim WafflePos

Sub WaffleShake()
    WafflePos = 3
    waffletime.Enabled = True
End Sub

Sub waffletime_Timer()
    waffle1.RotY = WafflePos
	If WafflePos <= 0.1 AND WafflePos >= -0.1 Then Me.Enabled = False:Exit Sub
    If WafflePos < 0 Then
        WafflePos = ABS(WafflePos)- 0.1
    Else
        WafflePos = - WafflePos + 0.1
    End If
End Sub

Dim WafflePos2

Sub WaffleShake2()
    WafflePos2 = 3
    waffletime2.Enabled = True
End Sub

Sub waffletime2_Timer()
    waffle2.RotY = WafflePos2
    If WafflePos2 <= 0.1 AND WafflePos2 >= -0.1 Then Me.Enabled = False:Exit Sub
    If WafflePos2 < 0 Then
        WafflePos2 = ABS(WafflePos2)- 0.1
    Else
        WafflePos2 = - WafflePos2 + 0.1
    End If
End Sub

Dim WafflePos3

Sub WaffleShake3()
    WafflePos3 = 3
    waffletime3.Enabled = True
End Sub

Sub waffletime3_Timer()
    waffle3.RotY = WafflePos3
    If WafflePos3 <= 0.1 AND WafflePos3 >= -0.1 Then Me.Enabled = False:Exit Sub
    If WafflePos3 < 0 Then
        WafflePos3 = ABS(WafflePos3)- 0.1
    Else
        WafflePos3 = - WafflePos3 + 0.1
    End If
End Sub

Dim WafflePos4

Sub WaffleShake4()
    WafflePos4 = 3
    waffletime4.Enabled = True
End Sub

Sub waffletime4_Timer()
    waffle4.RotY = WafflePos4
    If WafflePos4 <= 0.1 AND WafflePos4 >= -0.1 Then Me.Enabled = False:Exit Sub
    If WafflePos4 < 0 Then
        WafflePos4 = ABS(WafflePos4)- 0.1
    Else
        WafflePos4 = - WafflePos4 + 0.1
    End If
End Sub

'*****************
'  ESCAPE Targets
'*****************

Sub escape1_Hit
	startB2S(2)
	startB2S(1)
	startB2S(3)
	DOF 135, DOFPulse
    PlaySound SoundFX("fx_target",DOFTargets), 0, 1, pan(ActiveBall)
	PlaySound "runtarget"
    If Tilted Then Exit Sub
	If le2.State = 0 Then
	le2.State = 1
	le1.State = 0
	Exit Sub
	End If
	If le2.State = 1 Then
	le2.State = 1
	le1.State = 1
	End If
    AddScore 25010
    ' Do some sound or light effect
    FlashForMs f1, 1000, 50, 0
    LastSwitchHit = "escape1"
    ' do some check
    CheckESCAPETargets
End Sub

Sub escape2_Hit
	startB2S(2)
	startB2S(1)
	startB2S(3)
	DOF 134, DOFPulse
    PlaySound "fx_target", 0, 1, pan(ActiveBall)
	PlaySound "runtarget"
    If Tilted Then Exit Sub
	If le4.State = 0 Then
	le4.State = 1
	le3.State = 0
	Exit Sub
	End If
	If le4.State = 1 Then
	le4.State = 1
	le3.State = 1
	End If
    AddScore 25010
    ' Do some sound or light effect
    FlashForMs f1, 1000, 50, 0
    LastSwitchHit = "escape2"
    ' do some check
    CheckESCAPETargets
End Sub

Sub escape3_Hit
	startB2S(2)
	startB2S(1)
	startB2S(3)
	DOF 135, DOFPulse
    PlaySound SoundFX("fx_target",DOFTargets), 0, 1, pan(ActiveBall)
	PlaySound "runtarget"
    If Tilted Then Exit Sub
	If le6.State = 0 Then
	le6.State = 1
	le5.State = 0
	Exit Sub
	End If
	If le6.State = 1 Then
	le6.State = 1
	le5.State = 1
	End If
    AddScore 25010
    ' Do some sound or light effect
    FlashForMs f1, 1000, 50, 0
    LastSwitchHit = "escape3"
    ' do some check
    CheckESCAPETargets
End Sub

Sub escape4_Hit
	startB2S(2)
	startB2S(1)
	startB2S(3)
	DOF 135, DOFPulse
    PlaySound SoundFX("fx_target",DOFTargets), 0, 1, pan(ActiveBall)
	PlaySound "runtarget"
    If Tilted Then Exit Sub
	If le11.State = 0 Then
	le11.State = 1
	le10.State = 0
	Exit Sub
	End If
	If le11.State = 1 Then
	le11.State = 1
	le10.State = 1
	End If
    AddScore 25010
    ' Do some sound or light effect
    FlashForMs f1, 1000, 50, 0
    LastSwitchHit = "escape4"
    ' do some check
    CheckESCAPETargets
End Sub

Sub escape5_Hit
	startB2S(2)
	startB2S(1)
	startB2S(3)
	DOF 135, DOFPulse
    PlaySound SoundFX("fx_target",DOFTargets), 0, 1, pan(ActiveBall)
	PlaySound "runtarget"
    If Tilted Then Exit Sub
	If le13.State = 0 Then
	le13.State = 1
	le12.State = 0
	Exit Sub
	End If
	If le13.State = 1 Then
	le13.State = 1
	le12.State = 1
	End If
    AddScore 25010
    ' Do some sound or light effect
    FlashForMs f1, 1000, 50, 0
    LastSwitchHit = "escape5"
    ' do some check
    CheckESCAPETargets
End Sub

Sub escape6_Hit
	startB2S(2)
	startB2S(1)
	startB2S(3)
	DOF 135, DOFPulse
    PlaySound SoundFX("fx_target",DOFTargets), 0, 1, pan(ActiveBall)
	PlaySound "runtarget"
    If Tilted Then Exit Sub
	If le15.State = 0 Then
	le15.State = 1
	le14.State = 0
	Exit Sub
	End If
	If le15.State = 1 Then
	le15.State = 1
	le14.State = 1
	End If
    AddScore 25010
    ' Do some sound or light effect
    FlashForMs f1, 1000, 50, 0
    LastSwitchHit = "escape6"
    ' do some check
    CheckESCAPETargets
End Sub

Sub CheckESCAPETargets
If le1.State + le2.State + le3.State + le4.State + le5.State + le6.State + le10.State + le11.State + le12.State + le13.State + le14.State + le15.State = 12 Then
	Gate8.Open = True
	PlaySound "bell"
	MagnetU.MagnetON = True ' Magnet On
	le7.State = 2
	le8.State = 2
	le9.State = 2
End If
End Sub

Sub BallEscapeDrain_Hit
	MagnetU.MagnetON = False ' Magnet On
End Sub

Sub ResetESCAPELights
le1.State = 0
le2.State = 0
le3.State = 0
le4.State = 0
le5.State = 0
le6.State = 0
le7.State = 0
le8.State = 0
le9.State = 0	
le10.State = 0
le11.State = 0
le12.State = 0
le13.State = 0
le14.State = 0
le15.State = 0
End Sub

'***********
' Guards Targets or Chamber Targets - not sure what this art will end up as
'***********
Sub DropTargets
		If guard1.IsDropped = 0 or guard1.IsDropped = 0 or guard1.IsDropped = 0 Then DOF 122, DOFPulse
	guard1.IsDropped = 1
	guard2.IsDropped = 1
	guard3.IsDropped = 1
End Sub

Sub RaiseTargets
		If guard1.IsDropped = 1 or guard1.IsDropped = 1 or guard1.IsDropped = 1 Then DOF 122, DOFPulse
	guard1.IsDropped = 0
	guard2.IsDropped = 0
	guard3.IsDropped = 0
	If run1.IsDropped = 1 or run2.IsDropped = 1 or run3.IsDropped = 1 Then DOF 130, DOFPulse
	run1.IsDropped = 0
	run2.IsDropped = 0
	run3.IsDropped = 0
		If barb1.IsDropped = 1 or barb2.IsDropped = 1 or barb3.IsDropped = 1 or barb4.IsDropped = 1 Then DOF 129, DOFPulse
	barb1.IsDropped = 0
	barb2.IsDropped = 0
	barb3.IsDropped = 0
	barb4.IsDropped = 0
End Sub

Sub RaiseGuards
		If guard1.IsDropped = 1 or guard1.IsDropped = 1 or guard1.IsDropped = 1 Then DOF 122, DOFPulse
	guard1.IsDropped = 0
	guard2.IsDropped = 0
	guard3.IsDropped = 0
End Sub

Sub guard2_Hit
	startB2S(2)
	startB2S(1)
	startB2S(3)
	DOF 132, DOFPulse
	If WillMultiball = True Then
	AwardWill
	Else
	AddScore 20000
	PlaySound "portalopen"
	DMD "portal2-7.wmv", "", "",  8000   'Jackpot Bro
	CheckGUARDTargets
	End If
    LightEffect 2
	LightEffect 8
End Sub

Sub guard3_Hit
	startB2S(2)
	startB2S(1)
	startB2S(3)
	DOF 132, DOFPulse
	If WillMultiball = True Then
	AwardWill
	Else
	AddScore 20000
	PlaySound "portalopen"
	CheckGUARDTargets
	DMD "portal1-3.wmv", "", "",  4000   'Jackpot Bro
	End If
    LightEffect 2
	LightEffect 8
End Sub

Sub guard1_Hit
	startB2S(2)
	startB2S(1)
	startB2S(3)
	DOF 132, DOFPulse
	If WillMultiball = True Then
	AwardWill
	Else
	AddScore 20000
	PlaySound "portalopen"
	CheckGUARDTargets
	lc1.State = 2
	lc2.State = 2
	End If
	If bMultiBallMode = False Then
	Gate7.Open = True
	DMD "portal3-10.wmv", "", "",  11000   'Jackpot Bro
	DMD "black.png", "Guards Down", "Enter Upside Down",  1000
	End If
    LightEffect 2
	LightEffect 8
End Sub

Sub Gate7_Hit
If bMultiBallMode = False Then
Gate7.Open = True
DMD "black.png", "Guards Down", "Enter Upside Down",  1000
End If
End Sub

Sub CheckGUARDTargets
If WillMultiball = True Then Exit Sub
If guard1.IsDropped + guard2.IsDropped + guard3.IsDropped = 3 Then
DMD "black.png", "Guards Down", "Enter Upside Down",  1000
If bMultiBallMode = False Then
Gate7.Open = True
End If
lc1.State = 2
lc2.State = 2
End If
End Sub


'***********
' Spinner
'***********
' var spinvalue
' spinner increase its value by 1000 each time is hit
' max value is 10.000

Sub Spinner1_Spin
	DOF 131, DOFPulse
    PlaySound "fx_spinner", 0, 1, -0.05, 0.05
    If Not Tilted Then
        ' any light effect?
        ' any DMD display?
        AddScore spinvalue
    End If
End Sub

Sub SpinCounterLO_Hit
	startB2S(5)
    If Not Tilted Then
        If spinvalue <10000 Then
            spinvalue = spinvalue + 1000
        End If
    End If
End Sub

Sub Spinner2_Spin
	DOF 130, DOFPulse
    PlaySound "fx_spinner", 0, 1, -0.05, 0.05
    If Not Tilted Then
        ' any light effect?
        ' any DMD display?
        AddScore spinvalue
    End If
End Sub

Sub SpinCounterRO_Hit
	startB2S(6)
    If Not Tilted Then
        If spinvalue <10000 Then
            spinvalue = spinvalue + 1000
        End If
    End If
End Sub

'*****************
'   The Orbits
'*****************

Sub RightLODone_Hit
	startB2S(7)
	OrbitAward
	If bMultiBallMode = False Then
		PlaySound "ro2"
	End If
    If Tilted Then Exit Sub
		If DemoMultiball = True Then
		KillMonster
		End If
    LastSwitchHit = "RightLODone"
End Sub

Sub RightIODone_Hit
	startB2S(7)
		PlaySound "lights"
		If bMultiBallMode = False Then
		AwardNancy
		End If
    If Tilted Then Exit Sub
		If BarbJackpots = True Then
		AwardBarb
		End If
		If DemoMultiball = True Then
		KillMonster
		End If
    LastSwitchHit = "RightIODone"
End Sub

Dim NancyHit(25)
Sub AwardNancy
	NancyHit(CurrentPlayer) = NancyHit(CurrentPlayer) + 1
	Select Case NancyHit(CurrentPlayer)
	Case 2
				nancyrandom
				AddScore 300000
	Case 6
				nancyrandom	
				AddScore 500000
	Case 9
				nancyrandom	
				AddScore 700000
	Case 12
				nancyrandom
				AddScore 900000	
	Case 16
				nancyrandom
				AddScore 1100000	
	Case 20
				nancyrandom
				AddScore 1300000	
	Case 24
				nancyrandom
				AddScore 1500000
				NancyHit(CurrentPlayer) = 0
	End Select
End Sub

Sub nancyrandom
		dim tmp
        tmp = INT(RND * 7)
        Select Case tmp
        Case 0:DMD "nancy1-2.wmv", "", "", 3000
        Case 1:DMD "nancy2-10.wmv", "", "", 11000
        Case 2:DMD "nancy3-4.wmv", "", "", 5000
        Case 3:DMD "nancy4-4.wmv", "", "", 5000
        Case 4:DMD "nancy5-2.wmv", "", "", 3000
        Case 5:DMD "nancy6-6.wmv", "", "", 7000
        Case 6:DMD "nancy7-2.wmv", "", "", 3000
    End Select
End Sub

Sub LeftODone_Hit
	startB2S(5)
	OrbitAward
	If bMultiBallMode = False Then
		PlaySound "lo1"
	End If
    If Tilted Then Exit Sub
		If DemoMultiball = True Then
		KillMonster
		End If
    LastSwitchHit = "LeftODone"
End Sub

Sub LeftIODone_Hit
	startB2S(7)
	PlaySound "steve"
    If Tilted Then Exit Sub
		If bMultiBallMode = False Then
		AwardSteve
		End If
		If BarbJackpots = True Then
		AwardBarb
		End If
		If DemoMultiball = True Then
		KillMonster
		End If
    LastSwitchHit = "LeftIODone"
End Sub

Dim SteveHit(25)
Sub AwardSteve
	SteveHit(CurrentPlayer) = SteveHit(CurrentPlayer) + 1
	Select Case SteveHit(CurrentPlayer)
	Case 2
				steverandom
				AddScore 300000
	Case 6
				steverandom
				AddScore 500000	
	Case 9
				steverandom
				AddScore 700000	
	Case 12
				steverandom	
				AddScore 900000
	Case 16
				steverandom
				AddScore 1100000	
	Case 20
				steverandom
				AddScore 1300000	
	Case 24
				steverandom
				AddScore 1500000
				SteveHit(CurrentPlayer) = 0
	End Select
End Sub

Sub steverandom
		dim tmp
        tmp = INT(RND * 7)
        Select Case tmp
        Case 0:DMD "steve1-7.wmv", "", "", 8000
        Case 1:DMD "steve2-2.wmv", "", "", 2000
        Case 2:DMD "steve3-4.wmv", "", "", 5000
        Case 3:DMD "steve4-12.wmv", "", "", 13000
        Case 4:DMD "steve5-3.wmv", "", "", 4000
        Case 5:DMD "steve6-2.wmv", "", "", 3000
        Case 6:DMD "steve7-2.wmv", "", "", 3000
    End Select
End Sub



dim orbithit(100)
Sub OrbitAward
	orbithit(CurrentPlayer) = orbithit(CurrentPlayer) + 1
	Select Case orbithit(CurrentPlayer)
            Case 1
				 DMD "black.png", "30 More Shots", "For Super Orbits",  1000
            Case 31
				 DMD "black.png", "Super Orbits", "1000000 Award",  1000
				AddScore 1000000
            Case 34
				 DMD "black.png", "30 More Shots", "For Double Super Orbits",  1000
            Case 64
				 DMD "black.png", "Double Super", "Orbits 2000000 Award",  1000
				AddScore 2000000
            Case 68
				 DMD "black.png", "30 More Shots", "For Ultra Orbits",  1000
            Case 98
				 DMD "black.png", "Ultra Orbits", "4000000 Award",  1000
				AddScore 4000000
				Orbithit(CurrentPlayer) = 0
	End Select
END Sub
'****************
'     Ramps
'****************

Sub LeftRampDone_Hit
	DOF 142, DOFPulse
	LightEffect 12
	lrflashnow
	startB2S(5)
	AwardLR 
    PlaySound "fx_metalrolling", 0, 1, pan(ActiveBall)
	PlaySound "ping"
    If Tilted Then Exit Sub
		If DemoMultiball = True Then
		KillMonster
		End If
	If RunMultiball = True Then
	AwardRun
	Else
	End If
	If llr1.State = 2 Then
	End If
    LastSwitchHit = "LeftRampDone"
End Sub

' Used to award Extra Balls


Sub AwardLR
	LRHits(CurrentPlayer) = LRHits(CurrentPlayer) +1
        Select Case LRHits(CurrentPlayer)
            Case 1
				 DMD "black.png", "11 More Shots", "For Extra Ball",  1000
            Case 2
				DMD "r1-7.wmv", "", "", 8000
            Case 3
            Case 4
				DMD "r2-11.wmv", "", "", 12000
            Case 5
				DMD "black.png", "7 More Shots", "For Extra Ball",  1000
            Case 9
				DMD "black.png", "3 More Shots", "For Extra Ball",  1000
            Case 12
				AwardExtraBall()
				DMD "nancy5-2.wmv", "", "",  3000
				LRHits(CurrentPlayer) = 0
        End Select

End Sub

Sub RightRampDone_Hit
	DOF 143, DOFPulse
	LightEffect 13
	startB2S(6)
	rrflashnow
	AwardRR 
    PlaySound "fx_metalrolling", 0, 1, pan(ActiveBall)
	PlaySound "portalopen"
    If Tilted Then Exit Sub
		If DemoMultiball = True Then
		KillMonster
		End If
	If RunMultiball = True Then
	AwardRun
	Else
	End If
	If lrr1.State = 2 Then

	End If
    LastSwitchHit = "RightRampDone"
End Sub

' Used to award Extra Balls
Sub AwardRR
	RRHits(CurrentPlayer) = RRHits(CurrentPlayer) +1
        Select Case RRHits(CurrentPlayer)
            Case 1
				 DMD "black.png", "6 More Shots", "For Extra Ball",  1000
            Case 2
				DMD "r4-5.wmv", "", "", 6000
            Case 3
				DMD "black.png", "4 More Shots", "For Extra Ball",  1000
            Case 4
            Case 7
				AwardExtraBall()
				DMD "nancy5-2.wmv", "", "",  3000
				RRHits(CurrentPlayer) = 0
        End Select

End Sub


'************************
'  Mode: Where's Barb, Run from the Bad Men, Save Will & Kill the Demogorgon
'************************
'****************
' Mode: Where's Barb (spoiler alert... she dead.)
'****************
' Light all BARB lights to lock Balls
' Lock 2 balls to light start multiball
' During multiball hit orbits to search for Barb - 6 shots 
' 

Sub StartBarb() 'Multiball
	DOF 125, DOFPulse
    startB2S(4)
	BallsInLock(CurrentPlayer) = 0
	bMultiBallMode = True
	barbMultiball = True
	AddMultiball 2
	EnableBallSaver 15
	BarbJackpots = True
	GiOff
	LightSeqaxmas.UpdateInterval = 150
	LightSeqaxmas.Play SeqRandom, 10, , 50000
    LightSeqFlasher.UpdateInterval = 150
    LightSeqFlasher.Play SeqRandom, 10, , 50000
	' turn up the lights and yell baby
	llo3.State = 2
	lro4.State = 2
End Sub

Sub BallLockBarbExit()
    BallLockBarb.Kick 90, 7
		DOF 119, DOFPulse		
		DOF 115, DOFPulse
	Gate4.Open = False
End Sub

Sub AwardBarb
	If LookForBarb = False Then
	BarbHits(CurrentPlayer) = BarbHits(CurrentPlayer) + 1
	End If
	Select Case BarbHits(CurrentPlayer)
        Case 1
			AddScore 2000000
			DMDFlush
			DMD "barb-j1-4.wmv", "", "", 5000
			PlaySound "demogorgon"
        Case 2
			AddScore 2000000
			DMD "barb-j2-13.wmv", "", "", 14000
			PlaySound "demogorgon"
        Case 3
			AddScore 2000000
			DMD "barb-j3-7.wmv", "", "", 8000
			PlaySound "demogorgon"
        Case 4
			AddScore 2000000
			DMD "barb-j4-4.wmv", "", "", 5000
			PlaySound "demogorgon"
        Case 5
			AddScore 2000000
			DMD "barb-j5-5.wmv", "", "", 6000
			PlaySound "demogorgon"
			FindBarb
        Case 6
			FindBarb
        Case 7
			FindBarb
        Case 8
			FindBarb
        Case 9
			FindBarb
        Case 10
			FindBarb
    End Select
End Sub

Sub	FindBarb
	BarbJackpots = False
	LookForBarb = True
	Gate4.Open = True
	MagnetB.MagnetON = True ' Magnet On
	llo1.State = 0
	lro2.State = 0
	lro3.State = 0
	llo2.State = 0
	lro4.State = 2
End Sub

Sub BarbSuper
	LookForBarb = False
	DMDFlush
	DMD "barb-end-6.wmv", "", "", 7000
	DMD "black.png", "Super Jackpot", "6 Million",  1000
	DMD "black.png", "Where's Barb", "Complete",  1000
	PlaySound "demogorgon"
	AddScore 6000000
	BarbHits(CurrentPlayer) = 0
	lm1.State = 1
	' turn up the lights and yell baby
	llo3.State = 2
	lro4.State = 2
	BarbJackpots = True
	BallLockBarbExit
End Sub 

Sub EndBarb()
	LookForBarb = False
    barbMultiball = False
	BarbJackpots = False
	llo1.State = 0
	llo3.State = 0
	lro2.State = 0
	lro4.State = 0
	CheckMONSTER
	BarbUp
	RunUp
	GiOn
	LightSeqFlasher.StopPlay
	LightSeqaxmas.StopPlay
End Sub

'****************
' Mode: RUN - Escape the Bad Men
'****************
' Light all RUN lights to lock Balls
' Lock 3 balls to light start multiball
' During multiball hit ramps to escape - 6 shots 

Sub StartRun() 'Multiball
	DOF 125, DOFPulse
    startB2S(1)
	'DMD "black.png", "RUN RUN RUN", "Multiball", 1000
	'DMD "run-start-21.wmv", "","",  22000
	BallsInRunLock(CurrentPlayer) = 0
	bMultiBallMode = True
	RunMultiball = True
	AddMultiball 2
	EnableBallSaver 15
	GiOff
	LightSeqaxmas.UpdateInterval = 150
	LightSeqaxmas.Play SeqRandom, 10, , 50000
    LightSeqFlasher.UpdateInterval = 150
    LightSeqFlasher.Play SeqRandom, 10, , 50000
	' turn up the lights and yell baby
	lrr1.State = 2
	llr1.State = 2
End Sub

Sub BallLockRunExit()
    BallLockRun.Kick 90, 7
		DOF 117, DOFPulse		
		DOF 115, DOFPulse
	Gate6.Open = False
End Sub

Sub AwardRun
	If RunAway = False Then
	RunHits(CurrentPlayer) = RunHits(CurrentPlayer) + 1
	End If
	Select Case RunHits(CurrentPlayer)
        Case 1
			AddScore 2000000
			DMDFlush
			DMD "run-j1-3.wmv", "", "", 4000
			PlaySound "steve"
        Case 2
			AddScore 2000000
			DMD "run-j2-4.wmv", "", "", 5000
			PlaySound "steve"
        Case 3
			AddScore 2000000
			DMD "run-j3-7.wmv", "", "", 8000
			PlaySound "steve"
        Case 4
			AddScore 2000000
			DMD "run-j4-4.wmv", "", "", 4500
			PlaySound "steve"
        Case 5
			AddScore 2000000
			DMD "run-j5-11.wmv", "", "", 12000
			PlaySound "steve"
			TrytoRun
        Case 6
			TrytoRun
        Case 7
			TrytoRun
        Case 8
			TrytoRun
        Case 9
			TrytoRun
    End Select
End Sub

Sub	TrytoRun
	RunAway = True
	Gate6.Open = True
	llr1.State = 0
	lrr1.State = 0
	llo3.State = 2
End Sub

Sub RUNSuper
	RunAway = False
	AddScore 6000000
	DMDFlush
	DMD "run-end-16.wmv", "", "", 17000
	DMD "black.png", "Super Jackpot", "6 Million",  1000
	DMD "black.png", "RUN", "Complete",  1000
	PlaySound "steve"
	RunHits(CurrentPlayer) = 0
	Gate6.Open = False
	lm2.State = 1
	' turn up the lights and yell baby
	lrr1.State = 2
	llr1.State = 2
	llo3.State = 0
	BallLockRunExit
End Sub 

Sub EndRun()
	RunAway = False
    Dim lamp
    RunMultiball = False
	Gate6.Open = False
	lrr1.State = 0
	llr1.State = 0
	llo3.State = 0
	CheckMONSTER
	ResetESCAPELights
	BarbUp
	RunUp
	GiOn
	FlashEffect 0
	Flashxmas 0
	LightSeqFlasher.StopPlay
	LightSeqaxmas.StopPlay
End Sub


'****************
' Mode: Save Will - Go to the Upside Down and find Will
'****************

Sub BallLockEscape_Hit
	DMDFLush
	If WillSuperReady = True Then
	WillSJ
	End If
	If MonsterFinalBlow = True Then
	DemoSuper
	End If
	If bMultiBallMode = False Then
	EnterUpsideDown
	End If
End Sub

Sub EnterUpsideDown
	SolULFlipper 1
	SolULFlipper 0
	SolULFlipper 1
	SolULFlipper 0
	SolURFlipper 1
	SolURFlipper 0
	SolURFlipper 1
	SolURFlipper 0
	SolLFlipper 0
	SolRFlipper 0
	BallLockEscape.DestroyBall
	lowerflippersoff = False
	If le1.State + le2.State + le3.State + le4.State + le5.State + le6.State + le10.State + le11.State + le12.State + le13.State + le14.State + le15.State = 0 Then
	Dim waittime
    waittime = 14000
    vpmtimer.addtimer waittime, "UpsideDown'" 
	PlaySound "monster2"
	DMD "will_start-14.wmv", "", "",  14000   
	GiOff
	GiLowerOn
	Else
    waittime = 1000
    vpmtimer.addtimer waittime, "UpsideDown'"
	PlaySound "bell"
	DMD "black.png", "Entering", "Upside Down",  1000   
	GiOff
	GiLowerOn
	End If 
End Sub

Sub UpsideDown
	BallEscapeRelease.CreateBall
	BallEscapeRelease.Kick 90, 7
	DOF 120, DOFPulse
End Sub

Sub BallEscapeDrain_Hit
BallEscapeDrain.DestroyBall
BallEscapeExit.CreateBall
BallEscapeExit.Kick 90, 7
PlaySound SoundFXDOF("fx_Ballrel", 138, DOFPulse, DOFContactors), 0, 1, 0.1, 0.1
DOF 115, DOFPulse
GiOn
GiLowerOff
lowerflippersoff = True
End Sub

Sub BallEscape_Hit
ResetESCAPELights
PlaySound "demogorgon"
BallEscape.DestroyBall
BallEscapeExit.CreateBall
BallEscapeExit.Kick 90, 7
PlaySound SoundFXDOF("fx_Ballrel", 138, DOFPulse, DOFContactors), 0, 1, 0.1, 0.1
DOF 115, DOFPulse
Gate8.Open = False
MagnetU.MagnetON = False
le7.State = 0
le8.State = 0
le9.State = 0
SavedWill
GiLowerOff
lowerflippersoff = True
End Sub
 
Sub SavedWill() 'Multiball
	DOF 125, DOFPulse
    startB2S(3)
	DMDFlush
	DMD "black.png", "Saved Will", "Multiball", 1000
	DMD "will-end-34.wmv", "", "",  35000   'Jackpot Bro
	bMultiBallMode = True
	DOF 127, DOFPulse
	WillMultiball = True
	AddMultiball 3
	EnableBallSaver 15
	GiOff
	LightSeqaxmas.UpdateInterval = 150
	LightSeqaxmas.Play SeqRandom, 10, , 50000
    LightSeqFlasher.UpdateInterval = 150
    LightSeqFlasher.Play SeqRandom, 10, , 50000
	If guard1.IsDropped = 1 or guard1.IsDropped = 1 or guard1.IsDropped = 1 Then DOF 122, DOFPulse
	guard1.IsDropped = 0
	guard2.IsDropped = 0
	guard3.IsDropped = 0
	' turn up the lights and yell baby
	lc1.State = 2
End Sub

Sub AwardWill
	If WillSuper = False Then
	WillHits(CurrentPlayer) = WillHits(CurrentPlayer) + 1
	End If
	Select Case WillHits(CurrentPlayer)
        Case 1
			AddScore 4000000
			PlaySound "nancy"
			DMD "will1-4.wmv", "", "",  5000   'Jackpot Bro
        Case 2
			AddScore 4000000
			DMD "will2-30.wmv", "", "",  31000   'Jackpot Bro
			PlaySound "nancy"
        Case 3
			AddScore 4000000
			DMD "will3-31.wmv", "", "",  32000   'Jackpot Bro
			PlaySound "nancy"
			TryforSuper
    End Select
End Sub

Sub	TryforSuper
	WillHits(CurrentPlayer) = 0
	WillSuperReady = True
	Gate7.Open = True
	DropTargets
End Sub

Sub WillSJ
	PlaySound "nancy"
	BallLockEscape.DestroyBall
	BallEscapeExit.CreateBall
	BallEscapeExit.Kick 90, 7
	DOF 138, DOFPulse		
	DOF 115, DOFPulse
	AddScore 6000000
	DMDFlush
	DMD "will5-15.wmv", "", "",  16000   'Jackpot Bro
	DMD "black.png", "Super Jackpot", "6 Million",  1000
	DMD "black.png", "Will", "Complete",  1000
	WillHits(CurrentPlayer) = 0
	lm3.State = 1
	If guard1.IsDropped = 1 or guard1.IsDropped = 1 or guard1.IsDropped = 1 Then DOF 122, DOFPulse
	guard1.IsDropped = 0
	guard2.IsDropped = 0
	guard3.IsDropped = 0
	WillSuperReady = False
End Sub 

Sub EndWill()
    WillMultiball = False
	WillSuperReady = False
	lc1.State = 0
	lc2.State = 0
	Gate7.Open = False
	Gate8.Open = False
	CheckMONSTER
	BarbUp
	RunUp
	GiOn
	FlashEffect 0
	Flashxmas 0
	LightSeqFlasher.StopPlay
	LightSeqaxmas.StopPlay
End Sub


'****************
' Wizard Mode: Kill the Demogorgon
'****************

Sub CheckMONSTER
	If lm1.State + lm2.State + lm3.State = 3 Then
		Hit11
	End If
End Sub

Sub Hit11
    startB2S(2)
FlasherEL.opacity = 100
MagnetR.MagnetON = True ' Magnet On
DMD "black.png", "Give 011", "The Ball Now!",  4000   'Jackpot Bro
KickerEL.enabled = True
	GiOff
End Sub
Sub bell
	PlaySound "bell"
End Sub
Dim demodefeated
Sub KickerEL_hit
	DMDFlush
	finalflips = True
	Dim waittime
	Dim waittime2
	Dim waittime3
	Dim waittime4
	Dim waittime5
	If MonsterFinalBlow = True Then
	DemoSuper
	LightSeqaxmas.UpdateInterval = 150
	LightSeqaxmas.Play SeqRandom, 10, , 50000
    LightSeqFlasher.UpdateInterval = 150
    LightSeqFlasher.Play SeqRandom, 10, , 50000
	waittime = 106000
	waittime2 = 90500
	waittime3 = 94500
	waittime4 = 98500
	waittime5 = 102500
	lm4.State = 1
	PlaySong "m_wait" 	
	DMD "demo-end-90.wmv", "", "",  90500   'Jackpot Bro
	DMD "black.png", "You are", "The Champion",  4000   'Jackpot Bro
	DMD "black.png", "You will", "all be safe",  4000   'Jackpot Bro
	DMD "black.png", "", "or Will you",  4000   'Jackpot Bro
	DMD "black.png", "Champion Jackpot", "100 Million",  4000   'Jackpot Bro
    vpmtimer.addtimer waittime2, "bell '"
    vpmtimer.addtimer waittime3, "bell '"
    vpmtimer.addtimer waittime4, "bell '"
    vpmtimer.addtimer waittime5, "bell '"
    vpmtimer.addtimer waittime5, "CurrentSong '"
    vpmtimer.addtimer waittime, "DemoSuper '" 
    vpmtimer.addtimer waittime5, "kickit2 '"
    vpmtimer.addtimer waittime, "EndDemo '" 
	demodefeated = True
	MagnetR.MagnetON = False ' Magnet Off	
	Else
	waittime = 23000
	DMD "demo-start-21.wmv", "","",  22000
    vpmtimer.addtimer waittime, "kickit '" 
    vpmtimer.addtimer waittime, "StartMonster '" 
	MagnetR.MagnetON = False ' Magnet Off
	End If
	GiOff
End Sub

Sub kickit
	finalflips = False
	KickerEL.Kick -15, 50
	DOF 118, DOFPulse		
	DOF 115, DOFPulse
	KickerEL.enabled = False
End Sub

Sub kickit2
	finalflips = False
	DropTargets
	KickerEL.Kick 0, 50
	DOF 118, DOFPulse		
	DOF 115, DOFPulse
	KickerEL.enabled = False
End Sub


Sub StartMonster() 'Multiball
	bMultiBallMode = True
		DOF 127, DOFPulse
	DemoMultiball = True
	AddMultiball 5
	EnableBallSaver 30
	FlasherEL.opacity = 0
	DemoHits = 0
	FlashEffect 2
	LightSeqaxmas.UpdateInterval = 150
	LightSeqaxmas.Play SeqRandom, 10, , 65000
    LightSeqFlasher.UpdateInterval = 150
    LightSeqFlasher.Play SeqRandom, 10, , 65000
	' turn up the lights and yell baby
	    Dim a
    For each a in aLights
        a.State = 2
    Next
	StartRainbow aLights
End Sub

Sub KillMonster
	If MonsterFinalBlow = False Then
	DemoHits = DemoHits + 1
	End If
	Select Case DemoHits
        Case 1
			AddScore 4000000
			DMD "black.png", "Demogorgon", "Jackpot 1",  2000   'Jackpot Bro
			PlaySound "demohit"
        Case 2
			AddScore 4000000
			'DMD "black.png", "Demogorgon", "Jackpot 2",  2000   'Jackpot Bro
			PlaySound "demohit"
        Case 3
			AddScore 4000000
			DMD "demo-j1-7.wmv", "","",  8000
			PlaySound "demohit"
        Case 4
			AddScore 4000000
			'DMD "black.png", "Demogorgon", "Jackpot 4",  2000   'Jackpot Bro
			PlaySound "demohit"
        Case 5
			AddScore 4000000
			DMD "black.png", "Demogorgon", "Jackpot 5",  2000   'Jackpot Bro
			PlaySound "demohit"
        Case 6
			AddScore 4000000
			DMD "black.png", "Demogorgon", "Jackpot 6",  2000   'Jackpot Bro
			PlaySound "demohit"
        Case 7
			AddScore 4000000
			DMD "black.png", "Demogorgon", "Jackpot 7",  2000   'Jackpot Bro
			PlaySound "demohit"
        Case 8
			AddScore 4000000
			'DMD "black.png", "Demogorgon", "Jackpot 8",  2000   'Jackpot Bro
			PlaySound "demohit"
        Case 9
			AddScore 4000000
			DMD "demo-j2-22", "", "",  23000   'Jackpot Bro
			PlaySound "demohit"
        Case 10
			AddScore 4000000
			'DMD "black.png", "Demogorgon", "Jackpot 10",  2000   'Jackpot Bro
			PlaySound "demohit"
        Case 11
			AddScore 4000000
			DMD "black.png", "Demogorgon", "Jackpot 11",  2000   'Jackpot Bro
			PlaySound "demohit"
        Case 12
			AddScore 4000000
			DMD "black.png", "Demogorgon", "Jackpot 12",  2000   'Jackpot Bro
			PlaySound "demohit"
        Case 13
			AddScore 4000000
			DMD "black.png", "Demogorgon", "Jackpot 13",  2000   'Jackpot Bro
			PlaySound "demohit"
        Case 14
			AddScore 4000000
			DMD "black.png", "Demogorgon", "Jackpot 14",  2000   'Jackpot Bro
			PlaySound "demohit"
        Case 15
			AddScore 4000000
			'DMD "black.png", "Demogorgon", "Jackpot 15",  2000   'Jackpot Bro
			PlaySound "demohit"
        Case 16
			AddScore 4000000
			DMD "demo-j3-39.wmv", "", "",  39500   'Jackpot Bro
			PlaySound "demohit"
        Case 17
			AddScore 4000000
			'DMD "black.png", "Demogorgon", "Jackpot 17",  2000   'Jackpot Bro
			PlaySound "demohit"
        Case 18
			AddScore 4000000
			'DMD "black.png", "Demogorgon", "Jackpot 18",  2000   'Jackpot Bro
			PlaySound "demohit"
        Case 19
			AddScore 4000000
			DMD "black.png", "Demogorgon", "Jackpot 19",  2000   'Jackpot Bro
			PlaySound "demohit"
        Case 20
			AddScore 4000000
			DMD "black.png", "Demogorgon", "Jackpot 20",  2000   'Jackpot Bro
			PlaySound "demohit"
			TryForFinalBlow
    End Select
End Sub

Sub TryForFinalBlow
	MonsterFinalBlow = True
	TurnOffPlayfieldLights()
	FlasherEL.opacity = 100
	MagnetR.MagnetON = True ' Magnet On
	DMD "black.png", "Give 011", "The Ball Now!",  4000   'Jackpot Bro
	KickerEL.enabled = True
End Sub

Sub DemoSuper
	AddScore 10000000
	lm4.State = 1
	DemoHits = 0
	MonsterFinalBlow = False
End Sub


Sub EndDemo()
    Dim lamp
	MonsterFinalBlow = False
    DemoMultiball = False
	lm1.State = 0
	lm2.State = 0
	lm3.State = 0
	StopRainbow
	ResetAllLightsColor
	FlasherEL.opacity = 0
	BarbUp
	RunUp
	GiOn
	FlashEffect 0
	LightSeqFlasher.StopPlay
	LightSeqaxmas.StopPlay
		If guard1.IsDropped = 1 or guard1.IsDropped = 1 or guard1.IsDropped = 1 Then DOF 122, DOFPulse
	guard1.IsDropped = 0
	guard2.IsDropped = 0
	guard3.IsDropped = 0
	If demodefeated = True Then
	lm4.State = 1
	End If
End Sub
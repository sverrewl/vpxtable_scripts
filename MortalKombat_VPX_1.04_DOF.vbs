Option Explicit
Randomize

'DOF by Arngrim
On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Dim UltraDMD
Const cGameName = "mkii"
'************
'UltraDMD
'************
Dim curDir

Dim animatedBorder00

Dim UseUDMD

Const UltraDMD_VideoMode_Stretch = 0
Const UltraDMD_VideoMode_Top = 1
Const UltraDMD_VideoMode_Middle = 2
Const UltraDMD_VideoMode_Bottom = 3
Const UltraDMD_Animation_FadeIn = 0
Const UltraDMD_Animation_FadeOut = 1
Const UltraDMD_Animation_ZoomIn = 2
Const UltraDMD_Animation_ZoomOut = 3
Const UltraDMD_Animation_ScrollOffLeft = 4
Const UltraDMD_Animation_ScrollOffRight = 5
Const UltraDMD_Animation_ScrollOnLeft = 6
Const UltraDMD_Animation_ScrollOnRight = 7
Const UltraDMD_Animation_ScrollOffUp = 8
Const UltraDMD_Animation_ScrollOffDown = 9
Const UltraDMD_Animation_ScrollOnUp = 10
Const UltraDMD_Animation_ScrollOnDown = 11





Const UltraDMD_Animation_None = 14

Sub LoadUltraDMD
    Set UltraDMD = CreateObject("UltraDMD.DMDObject")
    UltraDMD.Init

    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    curDir = fso.GetAbsolutePathName(".")
    Set fso = nothing

    ' A Major version change indicates the version is no longer backward compatible
    If Not UltraDMD.GetMajorVersion = 1 Then
        MsgBox "Incompatible Version of UltraDMD found."
        Exit Sub
    End If

    'A Minor version change indicates new features that are all backward compatible
    If UltraDMD.GetMinorVersion < 0 Then
        MsgBox "Incompatible Version of UltraDMD found.  Please update to version 1.0 or newer."
        Exit Sub
    End If

    UltraDMD.SetProjectFolder curDir & "\Mortal Kombat II.UltraDMD"
    UltraDMD.SetVideoStretchMode UltraDMD_VideoMode_Middle

End Sub

Sub DMD_DisplaySceneEx(bkgnd,toptext,topBrightness, topOutlineBrightness, bottomtext, bottomBrightness, bottomOutlineBrightness, animateIn,pauseTime,animateOut)
    If Not UltraDMD is Nothing Then
        Debug.Print bkgnd
        Debug.Print toptext
        Debug.Print topBrightness
        Debug.Print topOutlineBrightness
        Debug.Print bottomtext
        Debug.Print bottomBrightness
        Debug.Print bottomOutlineBrightness
        Debug.Print animateIn
        Debug.Print pauseTime
        Debug.Print animateOut
        UltraDMD.DisplayScene00Ex bkgnd, toptext, topBrightness, topOutlineBrightness, bottomtext, bottomBrightness, bottomOutlineBrightness, animateIn, pauseTime, animateOut
        If pauseTime > 0 OR animateIn < 14 OR animateOut < 14 Then
            Timer1.Enabled = True
        End If
    End If
End Sub


Sub DMDScene (background, toptext, topbright, bottomtext, bottombright, animatein, pause, animateout, prio)		'regular DMD call with priority
	If prio >= OldDMDPrio Then
		DMDSceneInt background, toptext, topbright, bottomtext, bottombright, animatein, pause, animateout
		OldDMDPrio = prio
	End If
End Sub


Sub AttractAnimUltraDND_Timer
    Me.enabled = 0
    ShowTableInfo
End Sub


Sub BallSaverAnimT_Timer
    DMD_DisplaySceneEx "DMD1.png", "", 14, 2, "BALL SAVED", -1, -1, UltraDMD_Animation_None, 100, UltraDMD_Animation_None  
    DMD_DisplaySceneEx "DMD1.png", "", 15, 4, "", -1, -1, UltraDMD_Animation_None, 100, UltraDMD_Animation_None
    DMD_DisplaySceneEx "DMD1.png", "", 14, 2, "BALL SAVED", -1, -1, UltraDMD_Animation_None, 100, UltraDMD_Animation_None  
    DMD_DisplaySceneEx "DMD1.png", "", 15, 4, "", -1, -1, UltraDMD_Animation_None, 100, UltraDMD_Animation_None
    DMD_DisplaySceneEx "DMD1.png", "", 14, 2, "BALL SAVED", -1, -1, UltraDMD_Animation_None, 100, UltraDMD_Animation_None  
    DMD_DisplaySceneEx "DMD1.png", "", 15, 4, "", -1, -1, UltraDMD_Animation_None, 100, UltraDMD_Animation_None
    DMD_DisplaySceneEx "DMD1.png", "", 14, 2, "BALL SAVED", -1, -1, UltraDMD_Animation_None, 100, UltraDMD_Animation_None  
    DMD_DisplaySceneEx "DMD1.png", "", 15, 4, "", -1, -1, UltraDMD_Animation_None, 100, UltraDMD_Animation_None
    DMD_DisplaySceneEx "DMD1.png", "", 14, 2, "BALL SAVED", -1, -1, UltraDMD_Animation_None, 100, UltraDMD_Animation_None  
    DMD_DisplaySceneEx "DMD1.png", "", 15, 4, "", -1, -1, UltraDMD_Animation_None, 100, UltraDMD_Animation_None
    BallSaverAnimT.enabled = 0
End Sub

Sub BallSaverAnim()
    BallSaverAnimT.enabled = 1
End sub

Sub JackpotT250_Timer
    DMD_DisplaySceneEx "DMD1.png", "", 14, 2, "JACKPOT 250000", -1, -1, UltraDMD_Animation_None, 100, UltraDMD_Animation_None  
    DMD_DisplaySceneEx "DMD1.png", "", 15, 4, "      ", -1, -1, UltraDMD_Animation_None, 100, UltraDMD_Animation_None
    DMD_DisplaySceneEx "DMD1.png", "", 14, 2, "JACKPOT 250000", -1, -1, UltraDMD_Animation_None, 100, UltraDMD_Animation_None  
    DMD_DisplaySceneEx "DMD1.png", "", 15, 4, "      ", -1, -1, UltraDMD_Animation_None, 100, UltraDMD_Animation_None
    DMD_DisplaySceneEx "DMD1.png", "", 14, 2, "JACKPOT 250000", -1, -1, UltraDMD_Animation_None, 100, UltraDMD_Animation_None  
    DMD_DisplaySceneEx "DMD1.png", "", 15, 4, "      ", -1, -1, UltraDMD_Animation_None, 100, UltraDMD_Animation_None
    DMD_DisplaySceneEx "DMD1.png", "", 14, 2, "JACKPOT 250000", -1, -1, UltraDMD_Animation_None, 100, UltraDMD_Animation_None 
    JackpotT250.enabled = 0
End Sub

Sub Jackpot250()
    JackpotT250.enabled = 1
End sub

Sub OrbitAninT_Timer
    DMD_DisplaySceneEx "DMD1.png", "", 14, 2, "ORBIT MAXIMUM", -1, -1, UltraDMD_Animation_None, 500, UltraDMD_Animation_None  
    DMD_DisplaySceneEx "DMD1.png", "", 15, 4, "    250000    ", -1, -1, UltraDMD_Animation_None, 500, UltraDMD_Animation_None
    DMD_DisplaySceneEx "DMD1.png", "", 14, 2, "ORBIT MAXIMUM", -1, -1, UltraDMD_Animation_None, 500, UltraDMD_Animation_None  
    DMD_DisplaySceneEx "DMD1.png", "", 15, 4, "    250000    ", -1, -1, UltraDMD_Animation_None, 500, UltraDMD_Animation_None
    DMD_DisplaySceneEx "DMD1.png", "", 14, 2, "ORBIT MAXIMUM", -1, -1, UltraDMD_Animation_None, 500, UltraDMD_Animation_None  
    DMD_DisplaySceneEx "DMD1.png", "", 15, 4, "    250000    ", -1, -1, UltraDMD_Animation_None, 500, UltraDMD_Animation_None
    OrbitAninT.enabled = 0
End Sub

Sub OrbitAnin()
    OrbitAninT.enabled = 1
End sub

Sub Targets250T_Timer
    DMD_DisplaySceneEx "DMD1.png", "", 14, 2, "TARGETS COMPLETE", -1, -1, UltraDMD_Animation_None, 500, UltraDMD_Animation_None  
    DMD_DisplaySceneEx "DMD1.png", "", 15, 4, "     250000     ", -1, -1, UltraDMD_Animation_None, 500, UltraDMD_Animation_None
    DMD_DisplaySceneEx "DMD1.png", "", 14, 2, "TARGETS COMPLETE", -1, -1, UltraDMD_Animation_None, 500, UltraDMD_Animation_None  
    DMD_DisplaySceneEx "DMD1.png", "", 15, 4, "     250000     ", -1, -1, UltraDMD_Animation_None, 500, UltraDMD_Animation_None
    DMD_DisplaySceneEx "DMD1.png", "", 14, 2, "TARGETS COMPLETE", -1, -1, UltraDMD_Animation_None, 500, UltraDMD_Animation_None  
    DMD_DisplaySceneEx "DMD1.png", "", 15, 4, "     250000     ", -1, -1, UltraDMD_Animation_None, 500, UltraDMD_Animation_None
    Targets250T.enabled = 0
End Sub

Sub TargetsAnim250()
    Targets250T.enabled = 1
End sub

Sub Targets150T_Timer
    DMD_DisplaySceneEx "DMD1.png", "", 14, 2, "TARGETS COMPLETE", -1, -1, UltraDMD_Animation_None, 500, UltraDMD_Animation_None  
    DMD_DisplaySceneEx "DMD1.png", "", 15, 4, "     150000     ", -1, -1, UltraDMD_Animation_None, 500, UltraDMD_Animation_None
    DMD_DisplaySceneEx "DMD1.png", "", 14, 2, "TARGETS COMPLETE", -1, -1, UltraDMD_Animation_None, 500, UltraDMD_Animation_None  
    DMD_DisplaySceneEx "DMD1.png", "", 15, 4, "     150000     ", -1, -1, UltraDMD_Animation_None, 500, UltraDMD_Animation_None
    DMD_DisplaySceneEx "DMD1.png", "", 14, 2, "TARGETS COMPLETE", -1, -1, UltraDMD_Animation_None, 500, UltraDMD_Animation_None  
    DMD_DisplaySceneEx "DMD1.png", "", 15, 4, "     150000     ", -1, -1, UltraDMD_Animation_None, 500, UltraDMD_Animation_None
    Targets150T.enabled = 0
End Sub

Sub TargetsAnim150()
    Targets150T.enabled = 1
End sub

Sub Targets50T_Timer
    DMD_DisplaySceneEx "DMD1.png", "", 14, 2, "TARGETS COMPLETE", -1, -1, UltraDMD_Animation_None, 500, UltraDMD_Animation_None  
    DMD_DisplaySceneEx "DMD1.png", "", 15, 4, "     50000     ", -1, -1, UltraDMD_Animation_None, 500, UltraDMD_Animation_None
    DMD_DisplaySceneEx "DMD1.png", "", 14, 2, "TARGETS COMPLETE", -1, -1, UltraDMD_Animation_None, 500, UltraDMD_Animation_None  
    DMD_DisplaySceneEx "DMD1.png", "", 15, 4, "     50000     ", -1, -1, UltraDMD_Animation_None, 500, UltraDMD_Animation_None
    DMD_DisplaySceneEx "DMD1.png", "", 14, 2, "TARGETS COMPLETE", -1, -1, UltraDMD_Animation_None, 500, UltraDMD_Animation_None  
    DMD_DisplaySceneEx "DMD1.png", "", 15, 4, "     50000     ", -1, -1, UltraDMD_Animation_None, 500, UltraDMD_Animation_None
    Targets50T.enabled = 0
End Sub

Sub TargetsAnim50()
    Targets50T.enabled = 1
End sub

Sub Targets10T_Timer
    DMD_DisplaySceneEx "DMD1.png", "", 14, 2, "TARGETS COMPLETE", -1, -1, UltraDMD_Animation_None, 500, UltraDMD_Animation_None  
    DMD_DisplaySceneEx "DMD1.png", "", 15, 4, "     10000     ", -1, -1, UltraDMD_Animation_None, 500, UltraDMD_Animation_None
    DMD_DisplaySceneEx "DMD1.png", "", 14, 2, "TARGETS COMPLETE", -1, -1, UltraDMD_Animation_None, 500, UltraDMD_Animation_None  
    DMD_DisplaySceneEx "DMD1.png", "", 15, 4, "     10000     ", -1, -1, UltraDMD_Animation_None, 500, UltraDMD_Animation_None
    DMD_DisplaySceneEx "DMD1.png", "", 14, 2, "TARGETS COMPLETE", -1, -1, UltraDMD_Animation_None, 500, UltraDMD_Animation_None  
    DMD_DisplaySceneEx "DMD1.png", "", 15, 4, "     10000     ", -1, -1, UltraDMD_Animation_None, 500, UltraDMD_Animation_None
    Targets10T.enabled = 0
End Sub

Sub TargetsAnim10()
    Targets10T.enabled = 1
End sub

   


Sub MultiballJackpotAnimT_Timer
    DMD_DisplaySceneEx "DMD1.png", "MULTIBALL", 14, 2, "SHOOT JACKPOTS", -1, -1, UltraDMD_Animation_None, 400, UltraDMD_Animation_None  
    DMD_DisplaySceneEx "DMD1.png", "", 15, 4, "", -1, -1, UltraDMD_Animation_None, 100, UltraDMD_Animation_None
    DMD_DisplaySceneEx "DMD1.png", "MULTIBALL", 14, 2, "SHOOT JACKPOTS", -1, -1, UltraDMD_Animation_None, 400, UltraDMD_Animation_None  
    DMD_DisplaySceneEx "DMD1.png", "", 15, 4, "", -1, -1, UltraDMD_Animation_None, 100, UltraDMD_Animation_None
    DMD_DisplaySceneEx "DMD1.png", "MULTIBALL", 14, 2, "SHOOT JACKPOTS", -1, -1, UltraDMD_Animation_None, 400, UltraDMD_Animation_None  
    DMD_DisplaySceneEx "DMD1.png", "", 15, 4, "", -1, -1, UltraDMD_Animation_None, 100, UltraDMD_Animation_None
    DMD_DisplaySceneEx "DMD1.png", "MULTIBALL", 14, 2, "SHOOT JACKPOTS", -1, -1, UltraDMD_Animation_None, 400, UltraDMD_Animation_None  
    DMD_DisplaySceneEx "DMD1.png", "", 15, 4, "", -1, -1, UltraDMD_Animation_None, 100, UltraDMD_Animation_None
    DMD_DisplaySceneEx "DMD1.png", "MULTIBALL", 14, 2, "SHOOT JACKPOTS", -1, -1, UltraDMD_Animation_None, 400, UltraDMD_Animation_None  
    DMD_DisplaySceneEx "DMD1.png", "", 15, 4, "", -1, -1, UltraDMD_Animation_None, 100, UltraDMD_Animation_None
    MultiballJackpotAnimT.enabled = 0
End Sub

Sub MultiballJackpotAnim()
    MultiballJackpotAnimT.enabled = 1
End sub


Sub TiltAnim()
    TiltT.enabled = 1
End Sub

Sub TiltT_Timer
    TiltT.enabled = 0
  If Tilted = True Then
    DMD_DisplaySceneEx "DMD1.png", "", 14, 2, "       TILT!", -1, -1, UltraDMD_Animation_None, 100, UltraDMD_Animation_None
    DMD_DisplaySceneEx "DMD1.png", "", 14, 2, "TILT!     ", -1, -1, UltraDMD_Animation_None, 100, UltraDMD_Animation_None
    TiltAnim()
  End If
End Sub
 
Sub StopAnimUDMD
Dim iScene
    UltraDMD.CancelRendering
  Select Case iScene
         Case 1 :   DMD_DisplaySceneEx "DMD1.png", "JAVIER", 15, 4, "PRESENT", -1, -1, UltraDMD_Animation_ScrollOnLeft, 5000, UltraDMD_Animation_ScrollOffRight
         Case 2 :   DMD_DisplaySceneEx "DMD1.png", "", 15, 4, "", -1, -1, UltraDMD_Animation_None, 500, UltraDMD_Animation_None
         Case 3 :   DMD_DisplaySceneEx "DMDMKII.png", "", 15, 4, "", -1, -1, UltraDMD_Animation_ScrollOnUp, 5000, UltraDMD_Animation_ScrollOffDown
         Case 4 :   DMD_DisplaySceneEx "DMD1.png", "", 15, 4, "", -1, -1, UltraDMD_Animation_None, 100, UltraDMD_Animation_None  
         Case 5 :   DMD_DisplaySceneEx "DMD1.png", "   Original Table   ", 15, 4, "    By Javier   ", -1, -1, UltraDMD_Animation_ScrollOnLeft, 3000, UltraDMD_Animation_ScrollOffRight
         Case 6 :   DMD_DisplaySceneEx "DMD1.png", "", 15, 4, "", -1, -1, UltraDMD_Animation_None, 500, UltraDMD_Animation_None  
         Case 7 :   DMD_DisplaySceneEx "DMD1.png", "", 14, 2, "       TILT!", -1, -1, UltraDMD_Animation_None, 100, UltraDMD_Animation_None
         Case 8 :   DMD_DisplaySceneEx "DMD1.png", "", 14, 2, "TILT!     ", -1, -1, UltraDMD_Animation_None, 100, UltraDMD_Animation_None 
  End Select
 
    iScene = (iScene + 1) MOD 8


End Sub


'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'----------------------------------------------------------------------------------------------------------------
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////



Dim Ballsize,BallMass
Ballsize = 50
BallMass = (Ballsize^3)/125000

' Load the core.vbs for supporting Subs and functions
LoadCoreVBS

Sub LoadCoreVBS
    On Error Resume Next
    ExecuteGlobal GetTextFile("core.vbs")
    If Err Then MsgBox "Can't open core.vbs"
    On Error Goto 0
End Sub


if MortalKombat.ShowDT Then
   CabinetL.visible = 1
   CabinetR.visible = 1
   Ramp15.visible = 1
   Ramp16.visible = 1
   backrail.visible = 1
 Else
   CabinetL.visible = 0
   CabinetR.visible = 0
   Ramp15.visible = 0
   Ramp16.visible = 0
   backrail.visible = 0
End If

' Define any Constants
Const TableName = "Mortal Kombat II"
Const MyVersion = "1.04"
Const MaxPlayers = 2
Const BallSaverTime = 10 'in seconds
Const MaxMultiplier = 99 'no limit in this game
Const BallsPerGame = 3   ' 3 or 5

' Define Global Variables
Dim PlayersPlayingGame
Dim CurrentPlayer
Dim Credits
Dim Bonus
Dim BonusPoints(4)
Dim BonusMultiplier(4)
Dim BallsRemaining(4)
Dim ExtraBallsAwards(4)
Dim Score(4)
Dim HighScore(4)
Dim HighScoreName(4)
Dim Jackpot
Dim Tilt
Dim TiltSensitivity
Dim Tilted
Dim TotalGamesPlayed
Dim mBalls2Eject
Dim SkillshotValue

' Define Game Control Variables
Dim LastSwitchHit
Dim BallsOnPlayfield
Dim BallsInLock
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
Dim bMusicOff
Dim bAutoPlunger
Dim bSkillshotReady
Dim bExtraBallWonThisBall
Dim bSuperJackpotMode
Dim plungerIM 'used mostly as an autofire plunger
Dim ttable
Dim characterLevel
Dim CurrentSongPause: CurrentSongPause = (CurrentPlayer)
Dim CharacterMode


Sub startB2S(aB2S)
	If B2SOn Then
		Controller.B2SSetData 1,0
		Controller.B2SSetData 2,0
		Controller.B2SSetData 3,0
		Controller.B2SSetData 4,0
		Controller.B2SSetData 5,0
		Controller.B2SSetData 6,0
		Controller.B2SSetData 7,0
		Controller.B2SSetData 8,0
		Controller.B2SSetData 9,0
		Controller.B2SSetData 10,0
		Controller.B2SSetData 11,0
		Controller.B2SSetData 12,0
		Controller.B2SSetData 13,0
		Controller.B2SSetData 14,0
		Controller.B2SSetData 15,0
		Controller.B2SSetData 16,0
		Controller.B2SSetData 17,0
		Controller.B2SSetData 18,0
		Controller.B2SSetData aB2S,1
	End If
End Sub




' *********************************************************************
'                Visual Pinball Defined Script Events
' *********************************************************************
Dim LMAG, RMAG, UMAG
Sub MortalKombat_Init()
	Dim i
    Randomize
    LoadEM
    LoadUltraDMD

    'Impulse Plunger as autoplunger
    Const IMPowerSetting = 55 ' Plunger Power
    Const IMTime = 1.1        ' Time in seconds for Full Plunge
    Set plungerIM = New cvpmImpulseP
    With plungerIM
        .InitImpulseP swplunger, IMPowerSetting, IMTime
        .Random 1.5
        .InitExitSnd "fx_kicker", "fx_solenoid"
        .CreateEvents "plungerIM"
    End With

    'CaptiveBall
    CapKicker1.CreateBall
    CapKicker1.kick 0,1
    CapKicker1.enabled= 0

    ' Left Magnet
    Set LMAG = New cvpmMagnet
    With LMAG
        .InitMagnet LMagnet, 2
		.GrabCenter = 0
        .CreateEvents "LMAG"
    End With


    ' Right Magnet
    Set RMAG = New cvpmMagnet
    With RMAG
        .InitMagnet RMagnet, 2
		.GrabCenter = 0
        .CreateEvents "RMAG"
    End With

    ' Upper Magnet
    Set UMAG = New cvpmMagnet
    With UMAG
        .InitMagnet UMagnet, 2
		.GrabCenter = 0
        .CreateEvents "UMAG"
    End With


    ' Misc. VP table objects Initialisation, droptargets, animations...
    VPObjects_Init

    'load saved values, highscore, names, jackpot
    Loadhs

    'Init main variables
    For i = 1 To MaxPlayers
        Score(i) = 0
        BonusPoints(i) = 0
        BonusMultiplier(i) = 1
        BallsRemaining(i) = BallsPerGame
        ExtraBallsAwards(i) = 0
    Next

    ' initalise the DMD display
    DMD_Init

    ' freeplay or coins
    bFreePlay = False 'we want coins

    ' initialse any other flags
    bOnTheFirstBall = False
    bBallInPlungerLane = False
    bBallSaverActive = False
    bBallSaverReady = False
    bMultiBallMode = False
    bGameInPlay = False
    bMusicOn = True
    bMusicOff = False
    bAutoPlunger = False
    BallsOnPlayfield = 0
    BallsInLock = 0
    BallsInHole = 0
    LastSwitchHit = ""
    Tilt = 0
    TiltSensitivity = 6
    Tilted = False
    EndOfGame()
    IntroAnim()
    HologramRampF.visible = 0
    CharactersResetLight()
End Sub

Sub Magnets_Ini
  If bMultiBallMode = True Then
	LMAG.MagnetOn = 1
'	PlaySound SoundFX("fx_magnet",DOFShaker)
	RMAG.MagnetOn = 1
'	PlaySound SoundFX("fx_magnet",DOFShaker)
	UMAG.MagnetOn = 1
'	PlaySound SoundFX("fx_magnet",DOFShaker) 
	DOF 121, DOFOn
  Else
    LMAG.MagnetOn = 0
    RMAG.MagnetOn = 0
    UMAG.MagnetOn = 0
	DOF 121, DOFOff
  End If 
End Sub

'******
' Keys
'******

Sub MortalKombat_KeyDown(ByVal Keycode)
    If Keycode = AddCreditKey Then
        Credits = Credits + 1
		DOF 126, DOFOn
        If(Tilted = False) Then
            DMDFlush
            DMD "_", CenterLine(1, "CREDITS: " & Credits), 0, eNone, eNone, eNone, 500, True, "fx_coin"
            Playsound "fx_coin"
            StopAnimUDMD
            DMD_DisplaySceneEx "DMD1.png", "CREDITS: "& Credits, 15, 4, "PRESS START", -1, -1, UltraDMD_Animation_None, 1000, UltraDMD_Animation_None
            DMD_DisplaySceneEx "DMD1.png", "PRESS START", 15, 4, "   BUTTON  ", -1, -1, UltraDMD_Animation_None, 700, UltraDMD_Animation_None
            DMD_DisplaySceneEx "DMD1.png", ":  ", 15, 4, "", -1, -1, UltraDMD_Animation_None, 300, UltraDMD_Animation_None
            DMD_DisplaySceneEx "DMD1.png", "PRESS START", 15, 4, "   BUTTON   ", -1, -1, UltraDMD_Animation_None, 700, UltraDMD_Animation_None
            DMD_DisplaySceneEx "DMD1.png", ":  ", 15, 4, "", -1, -1, UltraDMD_Animation_None, 300, UltraDMD_Animation_None
            DMD_DisplaySceneEx "DMD1.png", "PRESS START", 15, 4, "   BUTTON   ", -1, -1, UltraDMD_Animation_None, 700, UltraDMD_Animation_None
            DMD_DisplaySceneEx "DMD1.png", ":  ", 15, 4, "", -1, -1, UltraDMD_Animation_None, 300, UltraDMD_Animation_None
            DMD_DisplaySceneEx "DMD1.png", "PRESS START", 15, 4, "   BUTTON   ", -1, -1, UltraDMD_Animation_None, 700, UltraDMD_Animation_None
            If NOT bGameInPlay Then ShowTableInfo
        End If
    End If

    If keycode = PlungerKey Then
        PlungerIM.AutoFire
    End If
    If bGameInPlay AND NOT Tilted Then

        If keycode = LeftTiltKey Then Nudge 90, 6:PlaySound "fx_nudge", 0, 1, -0.1, 0.25:CheckTilt
        If keycode = RightTiltKey Then Nudge 270, 6:PlaySound "fx_nudge", 0, 1, 0.1, 0.25:CheckTilt
        If keycode = CenterTiltKey Then Nudge 0, 7:PlaySound "fx_nudge", 0, 1, 1, 0.25:CheckTilt

        If keycode = LeftFlipperKey Then SolLFlipper 1
        If keycode = RightFlipperKey Then SolRFlipper 1

        If keycode = StartGameKey Then

            If((PlayersPlayingGame < MaxPlayers) AND(bOnTheFirstBall = True) ) Then

                If(bFreePlay = True) Then
                    PlayersPlayingGame = PlayersPlayingGame + 1
                    TotalGamesPlayed = TotalGamesPlayed + 1
                Else
                    If(Credits > 0) then
                        PlayersPlayingGame = PlayersPlayingGame + 1
                        TotalGamesPlayed = TotalGamesPlayed + 1
                        Credits = Credits - 1
						If Credits < 1 Then DOF 126, DOFOff
                    Else
                        ' Not Enough Credits to start a game.
                        DMDFlush
                        DMD CenterLine(0, "CREDITS " & Credits), CenterLine(1, "INSERT COIN"), 0, eNone, eBlink, eNone, 500, True, ""
                        StopAnimUDMD
                        DMD_DisplaySceneEx "DMD1.png", "CREDITS "& Credits, 15, 4, "INSERT COIN", -1, -1, UltraDMD_Animation_None, 700, UltraDMD_Animation_None
                        DMD_DisplaySceneEx "DMD1.png", "CREDITS "& Credits, 15, 4, " ", -1, -1, UltraDMD_Animation_None, 50, UltraDMD_Animation_None
                        DMD_DisplaySceneEx "DMD1.png", "CREDITS "& Credits, 15, 4, "INSERT COIN", -1, -1, UltraDMD_Animation_None, 700, UltraDMD_Animation_None
                        DMD_DisplaySceneEx "DMD1.png", "CREDITS "& Credits, 15, 4, " ", -1, -1, UltraDMD_Animation_None, 50, UltraDMD_Animation_None
                        DMD_DisplaySceneEx "DMD1.png", "CREDITS "& Credits, 15, 4, "INSERT COIN", -1, -1, UltraDMD_Animation_None, 700, UltraDMD_Animation_None
                        DMD_DisplaySceneEx "DMD1.png", "CREDITS "& Credits, 15, 4, " ", -1, -1, UltraDMD_Animation_None, 50, UltraDMD_Animation_None
                        DMD_DisplaySceneEx "DMD1.png", "CREDITS "& Credits, 15, 4, "INSERT COIN", -1, -1, UltraDMD_Animation_None, 700, UltraDMD_Animation_None
 
                    End If
                End If
            End If
        End If
        Else ' If (GameInPlay)

            If keycode = StartGameKey Then
                If(bFreePlay = True) Then
                    If(BallsOnPlayfield = 0) Then
                        ResetForNewGame()
                    End If
                Else
                    If(Credits > 0) Then
                        If(BallsOnPlayfield = 0) Then
                            Credits = Credits - 1
							If Credits < 1 Then DOF 126, DOFOff
                            ResetForNewGame()
                        End If
                    Else
                        ' Not Enough Credits to start a game.
                        DMDFlush
                        DMD CenterLine(0, "CREDITS " & Credits), CenterLine(1, "INSERT COIN"), 0, eNone, eBlink, eNone, 500, True, "Kahn_5"
                        ShowTableInfo
                        StopAnimUDMD
                        DMD_DisplaySceneEx "DMD1.png", "CREDITS "& Credits, 15, 4, "INSERT COIN", -1, -1, UltraDMD_Animation_None, 700, UltraDMD_Animation_None
                        DMD_DisplaySceneEx "DMD1.png", "CREDITS "& Credits, 15, 4, " ", -1, -1, UltraDMD_Animation_None, 50, UltraDMD_Animation_None
                        DMD_DisplaySceneEx "DMD1.png", "CREDITS "& Credits, 15, 4, "INSERT COIN", -1, -1, UltraDMD_Animation_None, 700, UltraDMD_Animation_None
                        DMD_DisplaySceneEx "DMD1.png", "CREDITS "& Credits, 15, 4, " ", -1, -1, UltraDMD_Animation_None, 50, UltraDMD_Animation_None
                        DMD_DisplaySceneEx "DMD1.png", "CREDITS "& Credits, 15, 4, "INSERT COIN", -1, -1, UltraDMD_Animation_None, 700, UltraDMD_Animation_None
                        DMD_DisplaySceneEx "DMD1.png", "CREDITS "& Credits, 15, 4, " ", -1, -1, UltraDMD_Animation_None, 50, UltraDMD_Animation_None
                        DMD_DisplaySceneEx "DMD1.png", "CREDITS "& Credits, 15, 4, "INSERT COIN", -1, -1, UltraDMD_Animation_None, 700, UltraDMD_Animation_None
 
                    End If
                End If
            End If
    End If ' If (GameInPlay)

    If hsbModeActive Then EnterHighScoreKey(keycode)

' Table specific
End Sub




Sub MortalKombat_KeyUp(ByVal keycode)
    If bGameInPLay AND NOT Tilted Then
        If keycode = LeftFlipperKey Then SolLFlipper 0
        If keycode = RightFlipperKey Then SolRFlipper 0
    End If
End Sub




'*************
' Pause Table
'*************

Sub MortalKombat_Paused
End Sub

Sub MortalKombat_unPaused
End Sub

Sub MortalKombat_Exit():
	Savehs
	Controller.Stop
	If Not UltraDMD is Nothing Then
		If UltraDMD.IsRendering Then
			UltraDMD.CancelRendering
		End If	
		UltraDMD = NULL 
	End If
End Sub

'********************
' Special JP Flippers
'********************

Sub SolLFlipper(Enabled)
    If Enabled Then
        PlaySound SoundFXDOF("fx_flipperup",101,DOFOn,DOFFlippers), 0, 1, -0.05, 0.15
        LeftFlipper.RotateToEnd
        RotateLaneLightsLeft
        RotateLaneLightsLeftUp
        
    Else
        PlaySound SoundFXDOF("fx_flipperdown",101,DOFOff,DOFFlippers), 0, 1, -0.05, 0.15
        LeftFlipper.RotateToStart
    End If
End Sub

Sub SolRFlipper(Enabled)
    If Enabled Then
        PlaySound SoundFXDOF("fx_flipperup",102,DOFOn,DOFFlippers), 0, 1, 0.05, 0.15
        RightFlipper.RotateToEnd
        RotateLaneLightsRight
        RotateLaneLightsRightUp
 '       StartSlotmachine()
       
    Else
        PlaySound SoundFXDOF("fx_flipperdown",102,DOFOff,DOFFlippers), 0, 1, 0.05, 0.15
        RightFlipper.RotateToStart
    End If
End Sub

' flippers hit Sound

Sub LeftFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 10, -0.1, 0.15
End Sub

Sub Rightflipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 10, 0.1, 0.15
End Sub

Sub RotateLaneLightsLeft
    Dim TempState
    TempState = l10.State
    l10.State = l14.State
    l14.State = l15.State
    l15.State = l17.State
    l17.State = TempState
End Sub

Sub RotateLaneLightsRight
    Dim TempState
    TempState = l17.State
    l17.State = l15.State
    l15.State = l14.State
    l14.State = l10.State
    l10.State = TempState
End Sub


Sub RotateLaneLightsLeftUp
    Dim TempState
    TempState = l11.State
    l11.State = l12.State
    l12.State = l13.State
    l13.State = TempState
End Sub

Sub RotateLaneLightsRightUp
    Dim TempState
    TempState = l13.State
    l13.State = l12.State
    l12.State = l11.State
    l11.State = TempState
End Sub



Sub UpdateFlipperLogo
    LFLogo.RotY = LeftFlipper.CurrentAngle
    RFlogo.RotY = RightFlipper.CurrentAngle
End Sub



'*********
' TILT
'*********

'NOTE: The TiltDecreaseTimer Subtracts .01 from the "Tilt" variable every round

Sub CheckTilt                                      'Called when table is nudged
    Tilt = Tilt + TiltSensitivity                  'Add to tilt count
    TiltDecreaseTimer.Enabled = True
    If(Tilt > TiltSensitivity) AND(Tilt < 15) Then 'show a warning
        DMD_DisplaySceneEx "DMD1.png", "", 14, 2, "    CAREFUL    ", -1, -1, UltraDMD_Animation_None, 400, UltraDMD_Animation_None 
        DMD "_", CenterLine(1, "CAREFUL!"), 0, eNone, eBlinkFast, eNone, 500, True, ""
    End if
    If Tilt > 15 Then 'If more that 15 then TILT the table
        Tilted = True
        'display Tilt
        TiltAnim()
        DMDFlush
        DMD CenterLine(0, "TILT!"), "", 0, eBlinkFast, eNone, eNone, 200, False, ""
        DisableTable True
        TiltRecoveryTimer.Enabled = True 'start the Tilt delay to check for all the balls to be drained
    End If
End Sub

Sub TiltDecreaseTimer_Timer
    ' DecreaseTilt
    If Tilt > 0 Then
        Tilt = Tilt - 0.1
    Else
        Me.Enabled = False
    End If
End Sub

Sub DisableTable(Enabled)
    If Enabled Then
        'turn off GI and turn off all the lights
        GiOff
        'Disable slings, bumpers etc
        LeftFlipper.RotateToStart
        RightFlipper.RotateToStart
		DOF 101, DOFOff
		DOF 102, DOFOff
        'Bumper1.Force = 0

        LeftSlingshot.Disabled = 1
        RightSlingshot.Disabled = 1
    Else
        'turn back on GI and the lights
        GiOn
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
        Me.Enabled = False
    End If
' else retry (checks again in another second or so)
End Sub

'********************
' Music as wav sounds
'********************

Dim Song:Song = ""

Sub ChangeSong
    If bMusicOn Then
        StopSound Song
        If bGameInPLay = False Then
            Song = "bgout_MKTrack-8" & ".mp3"
            PlayMusic Song
            Holostep2 = 1
            vpmtimer.AddTimer 3000, "IntroAnim '"
        
        Else
            song = "bgout_MKTrack-" &INT(7 * RND(1) ) + 1& ".mp3"
            PlayMusic Song
            IntroAnimOff()
            If bMultiballMode Then
                Song = "mortal-kombat"
                StopAllMusic
            Else
'                If bCatchemMode Then
'                    Song = "mu_catch"
'                Else
                    If hsbModeActive Then
                   '     Song = "Champion"
                    Else
                        Song = "mu_main"
                    End If
                End If
            End If
            PlaySound Song, -1, 0.1
        End If
'    End If
End Sub


Sub MortalKombat_MusicDone
    If bMusicOn Then
        ChangeSong
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
OldGiState = 0   'start witht the Gi off

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
        If UBound(tmp) = 1 Then
            GiOff
        Else
            Gion
        End If
    End If
End Sub

Sub GiOn
	DOF 127, DOFOn
    Dim bulb
    For each bulb in aGiLights
        bulb.State = 1
    Next
   	Playsound "fx_relay_on"
    Primitive58.image = "ShaoKahn_On"
    AyeFlasher1.visible = 1
    AyeFlasher2.visible = 1
End Sub

Sub GiOff
	DOF 127, DOFOff
    Dim bulb
    For each bulb in aGiLights
        bulb.State = 0
    Next
	Playsound "fx_relay_off"
    Primitive58.image = "ShaoKahn_Diff3"
    AyeFlasher1.visible = 0
    AyeFlasher2.visible = 0
End Sub

' GI & light sequence effects

Sub GiEffect(n)
    Select Case n
        Case 0 'all off
            LightSeqGi2.Play SeqAlloff
        Case 1 'all blink
            LightSeqGi2.UpdateInterval = 4
            LightSeqGi2.Play SeqBlinking, , 5, 100
        Case 2 'random
            LightSeqGi2.UpdateInterval = 10
            LightSeqGi2.Play SeqRandom, 5, , 1000
        Case 3 'upon
            LightSeqGi2.UpdateInterval = 4
            LightSeqGi2.Play SeqUpOn, 5, 1
        Case 4 ' left-right-left
            LightSeqGi2.UpdateInterval = 5
            LightSeqGi2.Play SeqLeftOn, 10, 1
            LightSeqGi2.UpdateInterval = 5
            LightSeqGi2.Play SeqRightOn, 10, 1
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
    End Select
End Sub

' Flasher Effects

Dim FEStep, FEffect
FEStep = 0
FEffect = 0

Sub FlashEffect(n)
    Select case n
        Case 1:FEStep = 0:FEffect = 1:FlashEffectTimer.Enabled = 1 'all blink
        Case 2:FEStep = 0:FEffect = 2:FlashEffectTimer.Enabled = 1 'random
        Case 3:FEStep = 0:FEffect = 3:FlashEffectTimer.Enabled = 1 'upon
        Case 4:FEStep = 0:FEffect = 4:FlashEffectTimer.Enabled = 1 'ordered random :)
    End Select
End Sub

Sub FlashEffectTimer_Timer()
    Select Case FEffect
        Case 1
            FlashForms BackFlasher1, 2000, 40, 0
            FlashForms BackFlasher2, 2000, 40, 0
            FlashForms BackFlasher3, 2000, 40, 0
            FlashForms BackFlasher4, 2000, 40, 0
            FlashForms BackFlasher5, 2000, 40, 0
            FlashForms BackFlasher6, 2000, 40, 0
			DOF 210, DOFPulse
            FlashForms FlasherLeftRamp, 2000, 40, 0:DOF 200, DOFPulse
            FlashForms FlasherRightRamp, 2000, 40, 0:DOF 203, DOFPulse
            FlashForms LJ1, 2000, 40, 0:FlashForms LJ4, 2000, 40, 0
            FlashForms l6f, 2000, 40, 0:FlashForms l6, 2000, 40, 0
            FlashForMs FirestationFlasher1, 2000, 40, 0:FlashForMs FirestationFlasher2, 2000, 40, 0
            FlashForMs l20, 2000, 40, 0:FlashForMs l20f, 2000, 40, 0
            FlashForMs l21, 2000, 40, 0:FlashForMs l21f, 2000, 40, 0
            FlashForMs l22, 2000, 40, 0:FlashForMs l22f, 2000, 40, 0
            FlashForMs l23, 2000, 40, 0:FlashForMs l23f, 2000, 40, 0
            FlashEffectTimer.Enabled = 0
        Case 2
            Select Case INT(RND * 14)
                Case 0:FlashForms FlasherLeftRamp, 1000, 40, 0:DOF 201, DOFPulse
                Case 1:FlashForms FlasherRightRamp, 1000, 40, 0:DOF 203, DOFPulse
                Case 2:FlashForms l5f, 1000, 40, 0:FlashForms l5, 1000, 40, 0
                Case 3:FlashForms l6f, 1000, 40, 0:FlashForms l6, 1000, 40, 0
                Case 10:FlashForms l20, 1000, 40, 0:FlashForms l20f, 1000, 40, 0:FlashForms l21, 1000, 40, 0:FlashForms l21f, 1000, 40, 0
                Case 11:FlashForms l22, 1000, 40, 0:FlashForms l22f, 1000, 40, 0:FlashForms l23, 1000, 40, 0:FlashForms l23f, 1000, 40, 0
                Case 12:FlashForMs BumperFlash, 400, 40, 0:FlashForMs BumperFlash1, 400, 40, 0
                Case 13:FlashForMs FirestationFlasher1, 1000, 40, 0:FlashForMs FirestationFlasher2, 1000, 40, 0
            End Select
            If FEStep = 20 then FlashEffectTimer.Enabled = 0
        Case 3
            Select case FEStep
                Case 0:FlashForms FlasherLeftRamp, 250, 40, 0:FlashForms FlasherRightRamp, 250, 40, 0:FlashForms StayPuftFlasher, 250, 40, 0:DOF 202, DOFPulse
                Case 1:FlashForms ectoFlasher, 250, 40, 0:FlashForms LibraryFlasher, 250, 40, 0:FlashForms BooksFlasher, 250, 40, 0
                Case 2:FlashForms ContainerFlasher, 250, 40, 0:FlashForms Protonflasher, 250, 40, 0
                Case 3:FlashForms l5f, 250, 40, 0:FlashForms l5, 250, 40, 0:FlashForms l6f, 250, 40, 0:FlashForms l6, 250, 40, 0
                Case 4:FlashForms FlasherLeftRamp, 250, 40, 0:FlashForms FlasherRightRamp, 250, 40, 0:FlashForms StayPuftFlasher, 250, 40, 0:DOF 202, DOFPulse
                Case 5:FlashForms ectoFlasher, 250, 40, 0:FlashForms LibraryFlasher, 250, 40, 0:FlashForms BooksFlasher, 250, 40, 0
                Case 6:FlashForms ContainerFlasher, 250, 40, 0:FlashForms Protonflasher, 250, 40, 0
                Case 7:FlashForms l5f, 250, 40, 0:FlashForms l5, 250, 40, 0:FlashForms l6f, 250, 40, 0:FlashForms l6, 250, 40, 0
                Case 8:FlashForms FlasherLeftRamp, 250, 40, 0:FlashForms FlasherRightRamp, 250, 40, 0:FlashForms StayPuftFlasher, 250, 40, 0:DOF 202, DOFPulse
                Case 9:FlashForms ectoFlasher, 250, 40, 0:FlashForms LibraryFlasher, 250, 40, 0:FlashForms BooksFlasher, 250, 40, 0
                Case 10:FlashForms ContainerFlasher, 250, 40, 0:FlashForms Protonflasher, 250, 40, 0
                Case 11:FlashForms l5f, 250, 40, 0:FlashForms l5, 250, 40, 0:FlashForms l6f, 250, 40, 0:FlashForms l6, 250, 40, 0
                Case 12:FlashEffectTimer.Enabled = 0
            End Select
        Case 4
    End Select
    FEStep = FEStep + 1
End Sub

Dim BEStep, BEffect
BEStep = 0
BEffect = 0

Sub BackFlashEffect(n)
    Select case n
        Case 1:BEStep = 0:BEffect = 1:BFlashEffectTimer.Enabled = 1 'all blink
        Case 2:BEStep = 0:BEffect = 2:BFlashEffectTimer.Enabled = 1 'random
        Case 3:BEStep = 0:BEffect = 3:BFlashEffectTimer.Enabled = 1 'Left>Right 3 times
        Case 4:BEStep = 0:BEffect = 4:BFlashEffectTimer.Enabled = 1 'Right>Left 3 times
    End Select
End Sub

Sub BFlashEffectTimer_Timer()
    Select Case BEffect
        Case 1
            FlashForms BackFlasher1, 1500, 40, 0
            FlashForms BackFlasher2, 1500, 40, 0
            FlashForms BackFlasher3, 1500, 40, 0
            FlashForms BackFlasher4, 1500, 40, 0
            FlashForms BackFlasher5, 1500, 40, 0
            FlashForms BackFlasher6, 1500, 40, 0
			DOF 211, DOFPulse
            BFlashEffectTimer.Enabled = 0
        Case 2
            Select Case INT(RND * 6)
                Case 0:FlashForms BackFlasher1, 500, 40, 0
                Case 1:FlashForms BackFlasher2, 500, 40, 0
                Case 2:FlashForms BackFlasher3, 500, 40, 0
                Case 3:FlashForms BackFlasher4, 500, 40, 0
                Case 4:FlashForms BackFlasher5, 500, 40, 0
                Case 5:FlashForms BackFlasher6, 500, 40, 0
            End Select
			DOF 212, DOFPulse
            If BEStep = 20 then BFlashEffectTimer.Enabled = 0
        Case 3
            Select case BEStep
                Case 0:FlashForms BackFlasher1, 200, 40, 0
                Case 1:FlashForms BackFlasher2, 200, 40, 0
                Case 2:FlashForms BackFlasher3, 200, 40, 0
                Case 3:FlashForms BackFlasher4, 200, 40, 0
                Case 4:FlashForms BackFlasher5, 200, 40, 0
                Case 5:FlashForms BackFlasher6, 200, 40, 0
                Case 6:FlashForms BackFlasher1, 200, 40, 0
                Case 7:FlashForms BackFlasher2, 200, 40, 0
                Case 8:FlashForms BackFlasher3, 200, 40, 0
                Case 9:FlashForms BackFlasher4, 200, 40, 0
                Case 10:FlashForms BackFlasher5, 200, 40, 0
                Case 11:FlashForms BackFlasher6, 200, 40, 0
                Case 12:FlashForms BackFlasher1, 200, 40, 0
                Case 13:FlashForms BackFlasher2, 200, 40, 0
                Case 14:FlashForms BackFlasher3, 200, 40, 0
                Case 15:FlashForms BackFlasher4, 200, 40, 0
                Case 16:FlashForms BackFlasher5, 200, 40, 0
                Case 17:FlashForms BackFlasher6, 200, 40, 0
                Case 18:BFlashEffectTimer.Enabled = 0
            End Select
			DOF 213, DOFPulse
        Case 4
            Select case BEStep
                Case 0:FlashForms BackFlasher6, 200, 40, 0
                Case 1:FlashForms BackFlasher5, 200, 40, 0
                Case 2:FlashForms BackFlasher4, 200, 40, 0
                Case 3:FlashForms BackFlasher3, 200, 40, 0
                Case 4:FlashForms BackFlasher2, 200, 40, 0
                Case 5:FlashForms BackFlasher1, 200, 40, 0
                Case 6:FlashForms BackFlasher6, 200, 40, 0
                Case 7:FlashForms BackFlasher5, 200, 40, 0
                Case 8:FlashForms BackFlasher4, 200, 40, 0
                Case 9:FlashForms BackFlasher3, 200, 40, 0
                Case 10:FlashForms BackFlasher2, 200, 40, 0
                Case 11:FlashForms BackFlasher1, 200, 40, 0
                Case 12:FlashForms BackFlasher6, 200, 40, 0
                Case 13:FlashForms BackFlasher5, 200, 40, 0
                Case 14:FlashForms BackFlasher4, 200, 40, 0
                Case 15:FlashForms BackFlasher3, 200, 40, 0
                Case 16:FlashForms BackFlasher2, 200, 40, 0
                Case 17:FlashForms BackFlasher1, 200, 40, 0
                Case 18:BFlashEffectTimer.Enabled = 0
            End Select
			DOF 213, DOFPulse
    End Select
    BEStep = BEStep + 1
End Sub

' *********************************************************************
'                      Supporting Ball & Sound Functions
' *********************************************************************

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
    Vol = Csng(BallVel(ball) ^2 / 2000)
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "table1" is the name of the table
    Dim tmp
    tmp = ball.x * 2 / MortalKombat.width-1
    If tmp > 0 Then
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

'*****************************************
'      JP's VP10 Rolling Sounds
'*****************************************

Const tnob = 20 ' total number of balls
Const lob = 1   'number of locked balls
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
    If UBound(BOT) = 3 Then Exit Sub 'there are always 4 balls on this table

    ' play the rolling sound for each ball
    For b = lob to UBound(BOT)
        If BallVel(BOT(b)) > 1 Then
            If BOT(b).z < 30 Then
                ballpitch = Pitch(BOT(b))
            Else
                ballpitch = Pitch(BOT(b)) * 100
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

'******************************
' Diverse Collection Hit Sounds
'******************************

Sub aMetals_Hit(idx):PlaySound "fx_MetalHit", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
Sub aRubber_Bands_Hit(idx):PlaySound "fx_rubber", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
Sub aRubbers_Hit(idx):PlaySound "fx_postrubber", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
Sub aRubber_Pins_Hit(idx):PlaySound "fx_postrubber", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
Sub aYellowPins_Hit(idx):PlaySound "fx_postrubber", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
Sub aPlastics_Hit(idx):PlaySound "fx_PlasticHit", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
Sub aGates_Hit(idx):PlaySound "fx_Gate", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
Sub aMaderas_Hit(idx):PlaySound "fx_Woodhit", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub



Sub PlaySoundEffect()
    Dim n
    n = "po_effect" & INT(RND * 36) + 1
    PlaySound n, 0, 1, pan(ActiveBall)
End Sub

Sub PlayMKSound()
    Dim n
    n = "po_fanfare" & INT(RND * 6) + 1
    PlaySound n
End Sub


' Ramp Soundss
Sub RHelp1_Hit()
    StopSound "fx_metalrolling"
    PlaySound "fx_ballrampdrop", 0, 1, pan(ActiveBall)
End Sub

Sub RHelp2_Hit()
    StopSound "fx_metalrolling"
    PlaySound "fx_ballrampdrop", 0, 1, pan(ActiveBall)
End Sub


 Sub RampRFX_Hit()
     PlaySound "ramp_enter", 0, 1, 0.05, 0.15
End Sub


 Sub RampLFX_Hit()
     PlaySound "ramp_enter", 0, 1, -0.05, 0.15
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
        BonusMultiplier(i) = 1
        BallsRemaining(i) = BallsPerGame
        ExtraBallsAwards(i) = 0
        Coins(i) = 0
    Next



    ' initialise any other flags
    bMultiBallMode = False
    Tilt = 0

    ' initialise Game variables
    Game_Init()

    ' you may wish to start some music, play a sound, do whatever at this point

    ' set up the start delay to handle any Start of Game Attract Sequence
    vpmtimer.addtimer 1000, "FirstBall '"
End Sub

' This is used to delay the start of a game to allow any attract sequence to
' complete.  When it expires it creates a ball for the player to start playing with

Sub FirstBall
    DMD CenterLine(1, ""), CenterLine(1, ""), 206, eNone, eNone, eNone, 3000, True, "kahn_Feel_The_Wrath_of_Shao_Kahn"
    DMD_DisplaySceneEx "DMD1.png", "Feel The Wrath", 15, 4, "of Shao Kahn", -1, -1, UltraDMD_Animation_None, 100, UltraDMD_Animation_None
    DMD_DisplaySceneEx "DMD1.png", "", 15, 4, "", -1, -1, UltraDMD_Animation_None, 3, UltraDMD_Animation_None
    DMD_DisplaySceneEx "DMD1.png", "Feel The Wrath", 15, 4, "of Shao Kahn", -1, -1, UltraDMD_Animation_None, 100, UltraDMD_Animation_None
    DMD_DisplaySceneEx "DMD1.png", "", 15, 4, "", -1, -1, UltraDMD_Animation_None, 3, UltraDMD_Animation_None
    DMD_DisplaySceneEx "DMD1.png", "Feel The Wrath", 15, 4, "of Shao Kahn", -1, -1, UltraDMD_Animation_None, 100, UltraDMD_Animation_None
    DMD_DisplaySceneEx "DMD1.png", "", 15, 4, "", -1, -1, UltraDMD_Animation_None, 3, UltraDMD_Animation_None
    DMD_DisplaySceneEx "DMD1.png", "Feel The Wrath", 15, 4, "of Shao Kahn", -1, -1, UltraDMD_Animation_None, 100, UltraDMD_Animation_None
    DMD_DisplaySceneEx "DMD1.png", "", 15, 4, "", -1, -1, UltraDMD_Animation_None, 3, UltraDMD_Animation_None
    DMD_DisplaySceneEx "DMD1.png", "Feel The Wrath", 15, 4, "of Shao Kahn", -1, -1, UltraDMD_Animation_None, 100, UltraDMD_Animation_None
    DMD_DisplaySceneEx "DMD1.png", "", 15, 4, "", -1, -1, UltraDMD_Animation_None, 3, UltraDMD_Animation_None
    DMD_DisplaySceneEx "DMD1.png", "Feel The Wrath", 15, 4, "of Shao Kahn", -1, -1, UltraDMD_Animation_None, 100, UltraDMD_Animation_None
    DMD_DisplaySceneEx "DMD1.png", "", 15, 4, "", -1, -1, UltraDMD_Animation_None, 3, UltraDMD_Animation_None
    DMD_DisplaySceneEx "DMD1.png", "Feel The Wrath", 15, 4, "of Shao Kahn", -1, -1, UltraDMD_Animation_None, 100, UltraDMD_Animation_None
    DMD_DisplaySceneEx "DMD1.png", "", 15, 4, "", -1, -1, UltraDMD_Animation_None, 3, UltraDMD_Animation_None
    DMD_DisplaySceneEx "DMD1.png", "Feel The Wrath", 15, 4, "of Shao Kahn", -1, -1, UltraDMD_Animation_None, 100, UltraDMD_Animation_None
    DMD_DisplaySceneEx "DMD1.png", "", 15, 4, "", -1, -1, UltraDMD_Animation_None, 3, UltraDMD_Animation_None
    DMD_DisplaySceneEx "DMD1.png", "Feel The Wrath", 15, 4, "of Shao Kahn", -1, -1, UltraDMD_Animation_None, 100, UltraDMD_Animation_None
    DMD_DisplaySceneEx "DMD1.png", "", 15, 4, "", -1, -1, UltraDMD_Animation_None, 3, UltraDMD_Animation_None
    DMD_DisplaySceneEx "DMD1.png", "Feel The Wrath", 15, 4, "of Shao Kahn", -1, -1, UltraDMD_Animation_None, 100, UltraDMD_Animation_None
    DMD_DisplaySceneEx "DMD1.png", "", 15, 4, "", -1, -1, UltraDMD_Animation_None, 3, UltraDMD_Animation_None
    DMD_DisplaySceneEx "DMD1.png", "Feel The Wrath", 15, 4, "of Shao Kahn", -1, -1, UltraDMD_Animation_None, 100, UltraDMD_Animation_None
    DMD_DisplaySceneEx "DMD1.png", "", 15, 4, "", -1, -1, UltraDMD_Animation_None, 3, UltraDMD_Animation_None
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

    ' set the current players bonus multiplier back down to 1X
    SetBonusMultiplier 1

    ' reset any drop targets, lights, game modes etc..
    'LightShootAgain.State = 0
    Bonus = 0
    bExtraBallWonThisBall = False
    ResetNewBallLights()

    'This is a new ball, so activate the ballsaver
    bBallSaverReady = True

    'and the skillshot
    'bSkillShotReady = True 'no skillshot in this game

    'Change the music ?

    'Reset any table specific
    TargetBonus = 0
    BumperBonus = 0
    CharacterBonus = 0

End Sub

' Create a new ball on the Playfield

Sub CreateNewBall()
    ' create a ball in the plunger lane kicker.
    BallRelease.CreateSizedball BallSize / 2
    ' There is a (or another) ball on the playfield
    BallsOnPlayfield = BallsOnPlayfield + 1
    ' kick it out..
    PlaySound SoundFXDOF("fx_Ballrel",122,DOFPulse,DOFContactors), 0, 1, 0.1, 0.1
    BallRelease.Kick 90, 4

    ' if there is 2 or more balls then set the multibal flag
    If BallsOnPlayfield > 1 Then
        bMultiBallMode = True
		DOF 132, DOFPulse
        ChangeSong
       ' MagnetsTimer.enabled = 1
            l19.state = 2
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
    bAutoPlunger = True

    ' wait if there is a ball in the plunger lane
    If bBallInPlungerLane Then
        Exit Sub
    Else
        CreateNewBall()
        mBalls2Eject = mBalls2Eject -1
        If mBalls2Eject = 0 Then 'if there are no more balls to eject then stop the timer
            Me.Enabled = False
        End If
    End If
End Sub

' The Player has lost his ball (there are no more balls on the playfield).
' Handle any bonus points awarded
'
Sub EndOfBall()
    Dim BonusDelayTime
    ' the first ball has been lost. From this point on no new players can join in
    bOnTheFirstBall = False

    ' only process any of this if the table is not tilted.  (the tilt recovery
    ' mechanism will handle any extra balls or end of game)
    If(Tilted = False) Then
        Dim AwardPoints, TotalBonus
        AwardPoints = 0:TotalBonus = 0

' add in any bonus points (multipled by the bonus multiplier)
'AwardPoints = BonusPoints(CurrentPlayer) * BonusMultiplier(CurrentPlayer)
'AddScore AwardPoints
'debug.print "Bonus Points = " & AwardPoints
'DMD "", CenterLine(1, "BONUS: " & BonusPoints(CurrentPlayer) & " X" & BonusMultiplier(CurrentPlayer) ), 0, eNone, eBlink, eNone, 1000, True, ""

'this table uses several bonus
        AwardPoints = TargetBonus * 1000
        TotalBonus = TotalBonus + AwardPoints
        DMD CenterLine(0, FormatScore(AwardPoints) ), CenterLine(1, "TARGET BONUS: " & TargetBonus), 0, eBlink, eNone, eNone, 500, False, ""
        DMD_DisplaySceneEx "DMD1.png", "Target: " &TargetBonus, 15, 4, " Bonus: " &AwardPoints, -1, -1, UltraDMD_Animation_None, 2000, UltraDMD_Animation_None  
        DMD_DisplaySceneEx "DMD1.png", "", 15, 4, "", -1, -1, UltraDMD_Animation_None, 5, UltraDMD_Animation_None


        AwardPoints = CharacterBonus * 15000
        TotalBonus = TotalBonus + AwardPoints
        DMD CenterLine(0, FormatScore(AwardPoints) ), CenterLine(1, "CHARACTER BONUS" & CharacterBonus), 0, eBlink, eNone, eNone, 500, False, ""
        DMD_DisplaySceneEx "DMD1.png", "Character: " &CharacterBonus, 15, 4, "Bonus:" &AwardPoints, -1, -1, UltraDMD_Animation_None, 2000, UltraDMD_Animation_None  
        DMD_DisplaySceneEx "DMD1.png", "", 15, 4, "", -1, -1, UltraDMD_Animation_None, 5, UltraDMD_Animation_None

'        AwardPoints = PokemonBonus * 50000
'        TotalBonus = TotalBonus + AwardPoints
'        DMD CenterLine(0, FormatScore(AwardPoints) ), CenterLine(1, "POKEMON BONUS: " & PokemonBonus), 0, eBlink, eNone, eNone, 500, False, ""

'        AwardPoints = EggBonus * 100000
'        TotalBonus = TotalBonus + AwardPoints
'        DMD CenterLine(0, FormatScore(AwardPoints) ), CenterLine(1, "EGG BONUS: " & EggBonus), 0, eBlink, eNone, eNone, 500, False, ""

        AwardPoints = BumperBonus * 100000
        TotalBonus = TotalBonus + AwardPoints
        DMD CenterLine(0, FormatScore(AwardPoints) ), CenterLine(1, "BUMPER BONUS: "), 0, eBlink, eNone, eNone, 500, False, ""
        DMD_DisplaySceneEx "DMD1.png", "Bumper:" , 15, 4, " Bonus: "  &AwardPoints, -1, -1, UltraDMD_Animation_None, 2000, UltraDMD_Animation_None  
        DMD_DisplaySceneEx "DMD1.png", "", 15, 4, "", -1, -1, UltraDMD_Animation_None, 5, UltraDMD_Animation_None

'        If Balls = BallsPerGame Then 'this is the last ball, add the coins left to the score
'            AwardPoints = Coins(CurrentPlayer) * 1000
'            TotalBonus = TotalBonus + AwardPoints
'            DMD CenterLine(0, FormatScore(AwardPoints) ), CenterLine(1, "COIN BONUS: " & Coins(CurrentPlayer) ), 0, eBlink, eNone, eNone, 500, False, ""
       ' End If

        DMD CenterLine(0, FormatScore(TotalBonus) ), CenterLine(1, "TOTAL BONUS " & " X" & BonusMultiplier(CurrentPlayer) ), 0, eBlinkFast, eNone, eNone, 1000, True, ""
        TotalBonus = TotalBonus * BonusMultiplier(CurrentPlayer)
        DMD_DisplaySceneEx "DMD1.png", "Multiplier", 15, 4, "X " & BonusMultiplier(CurrentPlayer), -1, -1, UltraDMD_Animation_None, 1000, UltraDMD_Animation_None  
        DMD_DisplaySceneEx "DMD1.png", "TOTAL BONUS", 15, 4, "" &TotalBonus, -1, -1, UltraDMD_Animation_None, 1500, UltraDMD_Animation_None
        DMD_DisplaySceneEx "DMD1.png", "", 15, 4, "", -1, -1, UltraDMD_Animation_None, 5, UltraDMD_Animation_None
        ' add a bit of a delay to allow for the bonus points to be shown & added up
        BonusDelayTime = 10000
        vpmtimer.addtimer BonusDelayTime, "Addscore TotalBonus '"
    Else
        'no bonus to count so move quickly to the next stage
        BonusDelayTime = 100
    End If
    ' start the end of ball timer which allows you to add a delay at this point
    vpmtimer.addtimer BonusDelayTime, "EndOfBall2 '"
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
        debug.print "Extra Ball"
   '     DMD "_", CenterLine(1, ("EXTRA BALL") ), "_", eNone, eBlink, eNone, 500, True, ""
        DMD_DisplaySceneEx "DMD1.png", "", 15, 4, "EXTRA BALL", -1, -1, UltraDMD_Animation_None, 2000, UltraDMD_Animation_None
        DMD "", "", 30, eNone, eNone, eBlinkFast, 2000, True, "Friendship_Again"
        AwardExtraBallBlinkLightOff()
        ' yep got to give it to them
        ExtraBallsAwards(CurrentPlayer) = ExtraBallsAwards(CurrentPlayer) - 1

        ' if no more EB's then turn off any shoot again light
        If(ExtraBallsAwards(CurrentPlayer) = 0) Then
            LightShootAgain.State = 0
        End If

        ' You may wish to do a bit of a song AND dance at this point

        ' Create a new ball in the shooters lane
        CreateNewBall()
    Else ' no extra balls

        BallsRemaining(CurrentPlayer) = BallsRemaining(CurrentPlayer) - 1

        ' was that the last ball ?
        If(BallsRemaining(CurrentPlayer) <= 0) Then
            debug.print "No More Balls, High Score Entry"

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
    debug.print "EndOfBall - Complete"
    ' are there multiple players playing this game ?
    If(PlayersPlayingGame > 1) Then
        ' then move to the next player
        NextPlayer = CurrentPlayer + 1
        ' are we going from the last player back to the first
        ' (ie say from player 4 back to player 1)
        If(NextPlayer > PlayersPlayingGame) Then
            NextPlayer = 1
        End If
          if NextPlayer = 1 Then
            PlaySound "Player1"
          End If
          if NextPlayer = 2 Then
            PlaySound "Player2"
          End If
    Else
        NextPlayer = CurrentPlayer
    End If

    debug.print "Next Player = " & NextPlayer

    ' is it the end of the game ? (all balls been lost for all players)
    If((BallsRemaining(CurrentPlayer) <= 0) AND(BallsRemaining(NextPlayer) <= 0) ) Then
        ' you may wish to do some sort of Point Match free game award here
        ' generally only done when not in free play mode
        ' set the machine into game over mode
        EndOfGame()

    ' you may wish to put a Game Over message on the

    Else

        ' set the next player
        CurrentPlayer = NextPlayer
        ' make sure the correct display is up to date
        AddScore 0

        ' reset the playfield for the new player (or new ball)
        ResetForNewPlayerBall()

        ' AND create a new ball
        CreateNewBall()
    End If
End Sub






' This function is called at the End of the Game, it should reset all
' Drop targets, AND eject any 'held' balls, start any attract sequences etc..

Sub EndOfGame()
    debug.print "End Of Game"
    bGameInPLay = False
    StopSound "Player1"
    ' just ended your game then play the end of game tune
    ChangeSong
    PlaySound "Finish_Him"
    ' ensure that the flippers are down
    SolLFlipper 0
    SolRFlipper 0

    ' terminate all modes - eject locked balls

    ' set any lights for the attract mode
    GiOff
    StartAttractMode 1
' you may wish to light any Game Over Light you may have
End Sub

Function Balls
    Dim tmp
    tmp = BallsPerGame - BallsRemaining(CurrentPlayer) + 1
    If tmp > BallsPerGame Then
        Balls = BallsPerGame
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
Sub Drain_Hit()
    ' Destroy the ball
      TiltT.enabled = 0
      StopAnimUDMD
    Drain.DestroyBall
    BallsOnPlayfield = BallsOnPlayfield - 1
    ' pretend to knock the ball into the ball storage mech
    PlaySound "fx_drain"
    DOF 123, DOFPulse
    ' if there is a game in progress AND it is not Tilted
    If(bGameInPLay = True) AND(Tilted = False) Then

        ' is the ball saver active,
        If(bBallSaverActive = True) Then

            ' yep, create a new ball in the shooters lane
            '   CreateNewBall
            ' we use the Addmultiball in case the multiballs are being ejected
            AddMultiball 1
            ' you may wish to put something on a display or play a sound at this point
            BallSaverAnim
            DMD "_", CenterLine(1, "BALL SAVED"), 0, eNone, eBlinkfast, eNone, 800, True, ""
        Else 

            ' cancel any multiball if on last ball (ie. lost all other balls)
            If(BallsOnPlayfield = 1) Then
                ' AND in a multi-ball??
                If(bMultiBallMode = True) then
                    ' not in multiball mode any more
                    bMultiBallMode = False
                    ChangeGi "white"
                    ' you may wish to change any music over at this point and
                    ' turn off any multiball specific lights
                    ResetJackpotLights
                    CheckResetMultiball
                  '  MagnetsTimer.enabled = 0
                        l19.state = 0
                    ChangeSong
                End If
            End If

            ' was that the last ball on the playfield
            If(BallsOnPlayfield = 0) Then
                 StopSound Song
                ' handle the end of ball (change player, high score entry etc..)
                EndOfBall()
                ' End Modes and timers
'                If bcoinfrenzy Then StopCoinFrenzyTimer_Timer
'                If bPikachuTargetMode Then PikachuTargetTimer_Timer
'                If bCharizardMode Then StopCharizardTimer_Timer
                If bRampBonus Then StopRampBonusTimer_Timer
         '       If bLoopBonus Then StopLoopBonusTimer_Timer
               ' ReduceBallType
                ChangeGi "white"
            End If
        End If
    End If
End Sub

' The Ball has rolled out of the Plunger Lane and it is pressing down the trigger in the shooters lane
' Check to see if a ball saver mechanism is needed and if so fire it up.

Sub swPlungerRest_Hit()
    debug.print "ball in plunger lane"
    ' some sound according to the ball position
    PlaySound "fx_sensor", 0, 1, 0.15, 0.25
    bBallInPlungerLane = True
    ' turn on Launch light is there is one
'    LaunchLight.State = 2
    ' if autoplunger is on then kick the ball
    If bAutoPlunger = True Then
        debug.print "autofire is on"
        PlungerIM.AutoFire
		DOF 131, DOFPulse
		DOF 125, DOFPulse
    If bBallSaverActive = False Then
        bAutoPlunger = False
    End If
    End If
    ' if there is a need for a ball saver, then start off a timer
    ' only start if it is ready, and it is currently not running, else it will reset the time period
    If(bBallSaverReady = True) AND(BallSaverTime <> 0) And(bBallSaverActive = False) Then
        EnableBallSaver BallSaverTime
    End If
    'Start the skillshot if ready
    If bSkillShotReady Then
        ResetSkillShotTimer.Interval = 1000 * 5 ' 5 seconds
        ResetSkillShotTimer.Enabled = True
  '      LightSeqSkillshot.Play SeqAllOff
  '      LightSeqSkillshotHit.Play SeqBlinking, , 5, 150
    'PlaySound a sound
    End If
    ' remember last trigger hit by the ball.
    LastSwitchHit = "swPlungerRest"
End Sub

' The ball is released from the plunger turn off some flags and check for skillshot

Sub swPlungerRest_UnHit()
    bBallInPlungerLane = False
    BackFlashEffect 1
    SetFlash 7, 2
	DOF 205, DOFPulse
	DOF 131, DOFPulse
	DOF 125, DOFPulse
End Sub

Sub EnableBallSaver(seconds)
    debug.print "Ballsaver started"
    ' set our game flag
    bBallSaverActive = True
    bBallSaverReady = False
    bAutoPlunger = True
    ' start the timer
    BallSaverTimer.Interval = 1000 * seconds
    BallSaverTimer.Enabled = True
    BallSaverSpeedUpTimer.Interval = 1000 * seconds -(1000 * seconds) / 3
    BallSaverSpeedUpTimer.Enabled = True
    ' if you have a ball saver light you might want to turn it on at this point (or make it flash)
    LightShootAgain.BlinkInterval = 160
    LightShootAgain.State = 2
End Sub

' The ball saver timer has expired.  Turn it off AND reset the game flag
'
Sub BallSaverTimer_Timer()
    debug.print "Ballsaver ended"
    BallSaverTimer.Enabled = False
    ' clear the flag
    bBallSaverActive = False
    bAutoPlunger = False
    ' if you have a ball saver light then turn it off at this point
    LightShootAgain.State = 0
End Sub

Sub BallSaverSpeedUpTimer_Timer()
    debug.print "Ballsaver Speed Up Light"
    BallSaverSpeedUpTimer.Enabled = False
    ' Speed up the blinking
    LightShootAgain.BlinkInterval = 80
    LightShootAgain.State = 2
End Sub


Sub DrainTrigger_hit
  If(bBallSaverActive = True) and (BallsOnPlayfield >= 1) Then
     PlaySound "Kahn_2"
   Else
     DMD "_", CenterLine(1, ""), 205, eNone, eNone, eNone, 800, True, ""
     vpmTimer.AddTimer 1, "DrainFX"  
  End If
End Sub

Dim Ka: Ka = Array("Kahn_1", "Kahn_2", "Kahn_3", "Kahn_4", "Kahn_5", "Kahn_6")
Sub DrainFX(dummy)
   Dim a
    a = INT(Ubound(Ka) * RND(6) )
    'debug.print a
    PlaySound Ka(a), 0, 1, 0.0
End Sub

 



' *********************************************************************
'                      Supporting Score Functions
' *********************************************************************

' Add points to the score AND update the score board
'
Sub AddScore(points)
    If(Tilted = False) Then
        ' add the points to the current players score variable
        Score(CurrentPlayer) = Score(CurrentPlayer) + points '* BallType
        ' update the score displays
        DMDScore
    End if

' you may wish to check to see if the player has gotten a replay
End Sub

' Add bonus to the bonuspoints AND update the score board
'
Sub AddBonus(points)
    If(Tilted = False) Then
        ' add the bonus to the current players bonus variable
        BonusPoints(CurrentPlayer) = BonusPoints(CurrentPlayer) + points
        ' update the score displays
        DMDScore
    End if

' you may wish to check to see if the player has gotten a replay
End Sub

Sub AddCoin(n)
    If(Tilted = False) Then
        ' add the coins to the current players coin variable
        Coins(CurrentPlayer) = Coins(CurrentPlayer) + n
        ' update the score displays
        DMDScore
    End if

    ' check if there is enough coins to enable the update ball
    If Coins(CurrentPlayer) > 249 Then
        BallUpdateLight.State = 2
    Else
        BallUpdateLight.State = 0
    End If
End Sub

' Add some points to the current Jackpot.
'
Sub AddJackpot(points)
    ' Jackpots only generally increment in multiball mode AND not tilted
    ' but this doesn't have to be the case
    If(Tilted = False) Then

        If(bMultiBallMode = True) Then
            Jackpot = Jackpot + points
        ' you may wish to limit the jackpot to a upper limit, ie..
        '	If (Jackpot >= 6000) Then
        '		Jackpot = 6000
        ' 	End if
        End if
    End if
End Sub

' Will increment the Bonus Multiplier to the next level
'
Sub IncrementBonusMultiplier()
    Dim NewBonusLevel

    ' if not at the maximum bonus level
    if(BonusMultiplier(CurrentPlayer) < MaxMultiplier) then
        ' then set it the next next one AND set the lights
        NewBonusLevel = BonusMultiplier(CurrentPlayer) + 1
        SetBonusMultiplier(NewBonusLevel)
    End if
End Sub

' Set the Bonus Multiplier to the specified level AND set any lights accordingly
'
Sub SetBonusMultiplier(Level)
    ' Set the multiplier to the specified level
    BonusMultiplier(CurrentPlayer) = Level

    ' If the multiplier is 1 then turn off all the bonus lights
    If(BonusMultiplier(CurrentPlayer) = 1) Then
        LightBonus2x.State = 0
        LightBonus3x.State = 0
        LightBigBonus.State = 0
    Else
        ' there is a bonus, turn on all the lights upto the current level
        If(BonusMultiplier(CurrentPlayer) >= 2) Then
            If(BonusMultiplier(CurrentPlayer) >= 2) Then
                LightBonus2x.state = 1
            End If
            If(BonusMultiplier(CurrentPlayer) >= 3) Then
                LightBonus3x.state = 1
            End If
            If(BonusMultiplier(CurrentPlayer) >= 4) Then
                LightBigBonus.state = 1
            End If
        End If
    ' etc..
    End If
End Sub

Sub IncrementBonus(Amount)
    Dim Value
    AddBonus Amount * 1000
    Bonus = Bonus + Amount
End Sub

Sub AwardExtraBall()
	DOF 124, DOFPulse
	DOF 125, DOFPulse
    If NOT bExtraBallWonThisBall Then
    '    DMDFlush
    '    DMD RightLine(0, ""), CenterLine(1, "_"), 4, eNone, eBlink, eNone, 3000, True, "WellDone"
    '    DMD "", "", 4, eNone, eNone, eBlinkFast, 3000, True, "WellDone"
        ExtraBallsAwards(CurrentPlayer) = ExtraBallsAwards(CurrentPlayer) + 1
        bExtraBallWonThisBall = True
        AwardExtraBallBlinkLightOn()
        GiEffect 1
        LightEffect 2
    END If
End Sub

sub AwardExtraBallBlinkLightOn()
    ExtraBallLight.BlinkInterval = 150
    ExtraBallLight.state=2
    AwardExtraBallBlinkLightOnTimer.Enabled = True
end sub

sub AwardExtraBallBlinkLightOnTimer_timer()
    ExtraBallLight.state=1
    me.enabled=0
end sub

sub AwardExtraBallBlinkLightOff()
    ExtraBallLight.BlinkInterval = 150
    ExtraBallLight.state=2
    ExtraBallOff.Enabled = True
end sub

sub ExtraBallOff_timer()
    ExtraBallLight.state=0
    me.enabled=0
end sub



Sub AwardJackpot()
	DMDFlush
    DMD RightLine(0, ""), CenterLine(1, ""), 6, eNone, eBlink, eNone, 1000, True, "Excellent"
    DMD_DisplaySceneEx "DMD1.png", "Jackpot", 14, 2, "100000" , -1, -1, UltraDMD_Animation_None, 800, UltraDMD_Animation_None
    AddScore 100000
    GiEffect 1
    LightEffect 2
    SetFlash 1, 2:DOF 206, DOFPulse
    SetFlash 2, 2:DOF 207, DOFPulse
    SetFlash 7, 2:DOF 205, DOFPulse
    SetFlash 8, 2:DOF 209, DOFPulse
    LightEffect 1
    BackFlashEffect 1
End Sub

Sub AwardSuperJackpot()
	DMDFlush
    DMD RightLine(0, ""), CenterLine(1, ""), 42, eNone, eBlink, eBlink, 800, True, "FlawlessVictory"
    DMD RightLine(1, ""), RightLine(1, ""), 37, eNone, eNone, eBlink, 800, True, "Fatality"
    DMD_DisplaySceneEx "DMD1.png", "Super Jackpot", 14, 2, "500000" , -1, -1, UltraDMD_Animation_None, 800, UltraDMD_Animation_None
    AddScore 500000
    LJ1.state = 2
    LJ2.state = 2
    LJ3.state = 2
    LJ4.state = 2
    CenterRampLight.state = 0
    RightRampLight.state = 0
    SetFlash 1, 2:DOF 206, DOFPulse
    SetFlash 2, 2:DOF 207, DOFPulse
    SetFlash 7, 2:DOF 205, DOFPulse
    SetFlash 8, 2:DOF 209, DOFPulse
    GiEffect 1
    LightEffect 1
    BackFlashEffect 3
    BackFlashEffect 4
    LightSeqSuperJackpot.StopPlay
    bSuperJackpotMode = False
End Sub

Sub JackpotShaoKahnDefeated()
	DMDFlush
    DMD RightLine(0, ""), CenterLine(1, ""), 42, eNone, eBlink, eBlink, 800, True, "FlawlessVictory"
    DMD RightLine(1, ""), RightLine(1, ""), 37, eNone, eNone, eBlink, 800, True, "Fatality"
    DMD CenterLine(0, FormatScore(Jackpot * characterLevel) ), Centerline(1, ("SHAO KAHN") ), 0, eBlinkFast, eBlinkFast, eNone, 1000, True, ""
    DMD CenterLine(0, FormatScore(Jackpot * characterLevel) ), Centerline(1, ("KAHN DEFEATED") ), 0, eBlinkFast, eBlinkFast, eNone, 1000, True, ""
    DMD_DisplaySceneEx "DMD1.png", "Super Jackpot", 14, 2, "250000" , -1, -1, UltraDMD_Animation_None, 1000, UltraDMD_Animation_None
    DMD_DisplaySceneEx "DMD1.png", "SHAO KAHN", 14, 2, "IS DEFEATED" , -1, -1, UltraDMD_Animation_None, 1000, UltraDMD_Animation_None
    AddScore 250000
    LJ1.state = 2
    LJ2.state = 2
    LJ3.state = 2
    LJ4.state = 2
    CenterRampLight.state = 0
    RightRampLight.state = 0
    SetFlash 1, 2:DOF 206, DOFPulse
    SetFlash 2, 2:DOF 207, DOFPulse
    SetFlash 7, 2:DOF 205, DOFPulse
    SetFlash 8, 2:DOF 209, DOFPulse
    GiEffect 1
    LightEffect 1
    BackFlashEffect 3
    BackFlashEffect 4
    LightSeqSuperJackpot.StopPlay
    bSuperJackpotMode = False
End Sub

Sub JackpotKintaroDefeated()
	DMDFlush
    DMD RightLine(0, ""), CenterLine(1, ""), 42, eNone, eBlink, eBlink, 800, True, "FlawlessVictory"
    DMD RightLine(1, ""), RightLine(1, ""), 37, eNone, eNone, eBlink, 800, True, "Fatality"
    DMD CenterLine(0, FormatScore(Jackpot * characterLevel) ), Centerline(1, ("KINTARO") ), 0, eBlinkFast, eBlinkFast, eNone, 1000, True, ""
    DMD CenterLine(0, FormatScore(Jackpot * characterLevel) ), Centerline(1, ("KINTARO DEFEATED") ), 0, eBlinkFast, eBlinkFast, eNone, 1000, True, ""
    DMD_DisplaySceneEx "DMD1.png", "Super Jackpot", 14, 2, "250000" , -1, -1, UltraDMD_Animation_None, 1000, UltraDMD_Animation_None
    DMD_DisplaySceneEx "DMD1.png", "KINTARO", 14, 2, "IS DEFEATED" , -1, -1, UltraDMD_Animation_None, 1000, UltraDMD_Animation_None
    AddScore 250000
    LJ1.state = 2
    LJ2.state = 2
    LJ3.state = 2
    LJ4.state = 2
    CenterRampLight.state = 0
    RightRampLight.state = 0
    SetFlash 1, 2:DOF 206, DOFPulse
    SetFlash 2, 2:DOF 207, DOFPulse
    SetFlash 7, 2:DOF 205, DOFPulse
    SetFlash 8, 2:DOF 209, DOFPulse
    GiEffect 1
    LightEffect 1
    BackFlashEffect 3
    BackFlashEffect 4
    LightSeqSuperJackpot.StopPlay
    bSuperJackpotMode = False
End Sub


Sub JackpotUnknownCharacterDefeated()
  If UMultiball = 1 Then
    DMDFlush
    DMD RightLine(0, ""), CenterLine(1, ""), 42, eNone, eBlink, eBlink, 800, True, "FlawlessVictory"
    DMD RightLine(1, ""), RightLine(1, ""), 37, eNone, eNone, eBlink, 800, True, "Fatality"
    DMD CenterLine(0, FormatScore(Jackpot * characterLevel) ), Centerline(1, ("SMOKE") ), 0, eBlinkFast, eBlinkFast, eNone, 1000, True, ""
    DMD CenterLine(0, FormatScore(Jackpot * characterLevel) ), Centerline(1, ("SMOKE DEFEATED") ), 0, eBlinkFast, eBlinkFast, eNone, 1000, True, ""
    DMD_DisplaySceneEx "DMD1.png", "Super Jackpot", 14, 2, "250000" , -1, -1, UltraDMD_Animation_None, 1000, UltraDMD_Animation_None
    DMD_DisplaySceneEx "DMD1.png", "SMOKE", 14, 2, "IS DEFEATED" , -1, -1, UltraDMD_Animation_None, 1000, UltraDMD_Animation_None
  End If

  If UMultiball = 2 Then
    DMDFlush
    DMD RightLine(0, ""), CenterLine(1, ""), 42, eNone, eBlink, eBlink, 800, True, "FlawlessVictory"
    DMD RightLine(1, ""), RightLine(1, ""), 37, eNone, eNone, eBlink, 800, True, "Fatality"
    DMD CenterLine(0, FormatScore(Jackpot * characterLevel) ), Centerline(1, ("JADE") ), 0, eBlinkFast, eBlinkFast, eNone, 1000, True, ""
    DMD CenterLine(0, FormatScore(Jackpot * characterLevel) ), Centerline(1, ("JADE DEFEATED") ), 0, eBlinkFast, eBlinkFast, eNone, 1000, True, ""
    DMD_DisplaySceneEx "DMD1.png", "Super Jackpot", 14, 2, "250000" , -1, -1, UltraDMD_Animation_None, 1000, UltraDMD_Animation_None
    DMD_DisplaySceneEx "DMD1.png", "JADE", 14, 2, "IS DEFEATED" , -1, -1, UltraDMD_Animation_None, 1000, UltraDMD_Animation_None
  End If
 
    AddScore 250000
    LJ1.state = 2
    LJ2.state = 2
    LJ3.state = 2
    LJ4.state = 2
    CenterRampLight.state = 0
    RightRampLight.state = 0
    SetFlash 1, 2:DOF 206, DOFPulse
    SetFlash 2, 2:DOF 207, DOFPulse
    SetFlash 7, 2:DOF 205, DOFPulse
    SetFlash 8, 2:DOF 209, DOFPulse
    GiEffect 1
    LightEffect 1
    BackFlashEffect 3
    BackFlashEffect 4
    LightSeqSuperJackpot.StopPlay
    bSuperJackpotMode = False
End Sub


Sub AwardSkillshot()
	DMD CenterLine(0, FormatScore(SkillshotValue) ), Centerline(1, ("SKILLSHOT") ), 0, eBlinkFast, eBlink, eNone, 1000, True, ""
    AddScore SkillshotValue
    ResetSkillShotTimer_Timer
End Sub


Sub ResetJackpotLights()
     LJ1.state = 0
     LJ2.state = 0
     LJ3.state = 0
     LJ4.state = 0
     LJ5.state = 0
     l18.state = 2 
     CenterRampLight.state = 0
     RightRampLight.state = 0 
     ChangeGi "White"
End Sub







'*****************************
'    Load / Save / Highscore
'*****************************

Sub Loadhs
    Dim x
    x = LoadValue(TableName, "HighScore1")
    If(x <> "")Then HighScore(0) = CDbl(x)Else HighScore(0) = 100000 End If
    x = LoadValue(TableName, "HighScore1Name")
    If(x <> "")Then HighScoreName(0) = x Else HighScoreName(0) = "AAA" End If
    x = LoadValue(TableName, "HighScore2")
    If(x <> "")then HighScore(1) = CDbl(x)Else HighScore(1) = 100000 End If
    x = LoadValue(TableName, "HighScore2Name")
    If(x <> "")then HighScoreName(1) = x Else HighScoreName(1) = "BBB" End If
    x = LoadValue(TableName, "HighScore3")
    If(x <> "")then HighScore(2) = CDbl(x)Else HighScore(2) = 100000 End If
    x = LoadValue(TableName, "HighScore3Name")
    If(x <> "")then HighScoreName(2) = x Else HighScoreName(2) = "CCC" End If
    x = LoadValue(TableName, "HighScore4")
    If(x <> "")then HighScore(3) = CDbl(x)Else HighScore(3) = 100000 End If
    x = LoadValue(TableName, "HighScore4Name")
    If(x <> "")then HighScoreName(3) = x Else HighScoreName(3) = "DDD" End If
    x = LoadValue(TableName, "Credits")
    If(x <> "")then Credits = CInt(x)Else Credits = 0 End If

    'x = LoadValue(TableName, "Jackpot")
    'If(x <> "") then Jackpot = CDbl(x) Else Jackpot = 200000 End If
    x = LoadValue(TableName, "TotalGamesPlayed")
    If(x <> "")then TotalGamesPlayed = CInt(x)Else TotalGamesPlayed = 0 End If
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
    If Score(2) > tmp Then tmp = Score(2)
    If Score(3) > tmp Then tmp = Score(3)
    If Score(4) > tmp Then tmp = Score(4)

    If tmp > HighScore(1) Then 'add 1 credit for beating the highscore
        Credits = Credits + 1
		DOF 126, DOFOn
    End If

    If tmp > HighScore(3) Then
        PlaySound SoundFXDOF("fx_knocker",124,DOFPulse,DOFKnocker)
		DOF 125, DOFPulse
        HighScore(3) = tmp
        'enter player's name
        HighScoreEntryInit()
    Else
        EndOfBallComplete()
    End If
End Sub

Sub HighScoreEntryInit()
    hsbModeActive = True
    ChangeSong
    hsLetterFlash = 0
    PlaySound "Champion"

    hsEnteredDigits(0) = " "
    hsEnteredDigits(1) = " "
    hsEnteredDigits(2) = " "
    hsCurrentDigit = 0

    hsValidLetters = " ABCDEFGHIJKLMNOPQRSTUVWXYZ'<>*+-/=\^0123456789`" ' ` is back arrow
    hsCurrentLetter = 1
    DMDFlush()
    HighScoreDisplayNameNow()

    HighScoreFlashTimer.Interval = 250
    HighScoreFlashTimer.Enabled = True
End Sub

Sub EnterHighScoreKey(keycode)
    If keycode = LeftFlipperKey Then
        playsound "fx_Previous"
        hsCurrentLetter = hsCurrentLetter - 1
        if(hsCurrentLetter = 0) then
            hsCurrentLetter = len(hsValidLetters)
        end if
        HighScoreDisplayNameNow()
    End If

    If keycode = RightFlipperKey Then
        playsound "fx_Next"
        hsCurrentLetter = hsCurrentLetter + 1
        if(hsCurrentLetter > len(hsValidLetters) ) then
            hsCurrentLetter = 1
        end if
        HighScoreDisplayNameNow()
    End If

    If keycode = PlungerKey Then
        if(mid(hsValidLetters, hsCurrentLetter, 1) <> "`") then
            playsound "fx_Enter"
            hsEnteredDigits(hsCurrentDigit) = mid(hsValidLetters, hsCurrentLetter, 1)
            hsCurrentDigit = hsCurrentDigit + 1
            if(hsCurrentDigit = 3) then
                HighScoreCommitName()
            else
                HighScoreDisplayNameNow()
            end if
        else
            playsound "fx_Esc"
            hsEnteredDigits(hsCurrentDigit) = " "
            if(hsCurrentDigit > 0) then
                hsCurrentDigit = hsCurrentDigit - 1
            end if
            HighScoreDisplayNameNow()
        end if
    end if
End Sub

Sub HighScoreDisplayNameNow()
    HighScoreFlashTimer.Enabled = False
    hsLetterFlash = 0
    HighScoreDisplayName()
    HighScoreFlashTimer.Enabled = True
End Sub

Sub HighScoreDisplayName()
    Dim i
    Dim TempTopStr
    Dim TempBotStr

    TempTopStr = "YOUR NAME:"
    dLine(0) = ExpandLine(TempTopStr, 0)
    DMDUpdate 0

    TempBotStr = "    > "
    if(hsCurrentDigit > 0) then TempBotStr = TempBotStr & hsEnteredDigits(0)
    if(hsCurrentDigit > 1) then TempBotStr = TempBotStr & hsEnteredDigits(1)
    if(hsCurrentDigit > 2) then TempBotStr = TempBotStr & hsEnteredDigits(2)

    if(hsCurrentDigit <> 3) then
        if(hsLetterFlash <> 0) then
            TempBotStr = TempBotStr & "_"
        else
            TempBotStr = TempBotStr & mid(hsValidLetters, hsCurrentLetter, 1)
        end if
    end if

    if(hsCurrentDigit < 1) then TempBotStr = TempBotStr & hsEnteredDigits(1)
    if(hsCurrentDigit < 2) then TempBotStr = TempBotStr & hsEnteredDigits(2)

    TempBotStr = TempBotStr & " <    "
    dLine(1) = ExpandLine(TempBotStr, 1)
    DMDUpdate 1
End Sub

Sub HighScoreFlashTimer_Timer()
    HighScoreFlashTimer.Enabled = False
    hsLetterFlash = hsLetterFlash + 1
    if(hsLetterFlash = 2) then hsLetterFlash = 0
    HighScoreDisplayName()
    HighScoreFlashTimer.Enabled = True
End Sub

Sub HighScoreCommitName()
    HighScoreFlashTimer.Enabled = False
    hsbModeActive = False
    ChangeSong
    hsEnteredName = hsEnteredDigits(0) & hsEnteredDigits(1) & hsEnteredDigits(2)
    if(hsEnteredName = "   ") then
        hsEnteredName = "YOU"
    end if

    HighScoreName(3) = hsEnteredName
    SortHighscore
    EndOfBallComplete()
End Sub

Sub SortHighscore
    Dim tmp, tmp2, i, j
    For i = 0 to 3
        For j = 0 to 2
            If HighScore(j) < HighScore(j + 1) Then
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

' *********************************************************************
'   JP's Reduced Display Driver Functions (based on script by Black)
' only 5 effects: none, scroll left, scroll right, blink and blinkfast
' 3 Lines, treats all 3 lines as text. 3rd line is just 1 character
' Example format:
' DMD "text1","text2",number, eNone, eNone, eNone, 250, True, "sound"
' Short names:
' dq = display queue
' de = display effect
' *********************************************************************

Const eNone = 0        ' Instantly displayed
Const eScrollLeft = 1  ' scroll on from the right
Const eScrollRight = 2 ' scroll on from the left
Const eBlink = 3       ' Blink (blinks for 'TimeOn')
Const eBlinkFast = 4   ' Blink (blinks for 'TimeOn') at user specified intervals (fast speed)

Const dqSize = 64

Dim dqHead
Dim dqTail
Dim deSpeed
Dim deBlinkSlowRate
Dim deBlinkFastRate

Dim dCharsPerLine(2)
Dim dLine(2)
Dim deCount(2)
Dim deCountEnd(2)
Dim deBlinkCycle(2)

Dim dqText(2, 64)
Dim dqEffect(2, 64)
Dim dqTimeOn(64)
Dim dqbFlush(64)
Dim dqSound(64)

Sub DMD_Init() 'default/startup values
    Dim i, j
    DMDFlush()
    deSpeed = 20
    deBlinkSlowRate = 5
    deBlinkFastRate = 2
    dCharsPerLine(0) = 16
    dCharsPerLine(1) = 20
    dCharsPerLine(2) = 1
    For i = 0 to 2
        dLine(i) = Space(dCharsPerLine(i) )
        deCount(i) = 0
        deCountEnd(i) = 0
        deBlinkCycle(i) = 0
        dqTimeOn(i) = 0
        dqbFlush(i) = True
        dqSound(i) = ""
    Next
    For i = 0 to 2
        For j = 0 to 64
            dqText(i, j) = ""
            dqEffect(i, j) = eNone
        Next
    Next
    DMD dLine(0), dLine(1), dLine(2), eNone, eNone, eNone, 25, True, ""
End Sub

Sub DMDFlush()
    Dim i
    DMDTimer.Enabled = False
    DMDEffectTimer.Enabled = False
    dqHead = 0
    dqTail = 0
    For i = 0 to 2
        deCount(i) = 0
        deCountEnd(i) = 0
        deBlinkCycle(i) = 0
    Next
End Sub


Sub DMDScore()
    Dim tmp, tmp1, tmp2
    if(dqHead = dqTail) Then
        tmp = RightLine(0, FormatScore(Score(Currentplayer) ) )
        'tmp = CenterLine(0, FormatScore(Score(Currentplayer) ) )
        tmp1 = CenterLine(1, "PLAYER" & CurrentPlayer & " j" &  "BALL" & Balls)' & Coins(Currentplayer) )
        UltraDMD.SetScoreboardBackgroundImage "DMD1.png", 15, 15
       ' UltraDMD.DisplayScoreboard 1, 1, "" & Score(1) , Score(2), Score(3), Score(4), "BALL "& Balls, "CREDITS:" & Credits
        UltraDMD.DisplayScoreboard PlayersPlayingGame, CurrentPlayer, Score(1), Score(2), Score(3), Score(4), "Ball " & Balls, "Credits: " & Credits
        'tmp1 = CenterLine(1, "PLAYER " & CurrentPlayer & " BALL " & Balls)
        'tmp1 = FormatScore(Bonuspoints(Currentplayer) ) & " X" &BonusMultiplier(Currentplayer)
        tmp2 = 0
    End If
    DMD tmp, tmp1, tmp2, eNone, eNone, eNone, 25, True, ""
End Sub
 
Sub DMD(Text0, Text1, Text2, Effect0, Effect1, Effect2, TimeOn, bFlush, Sound)
    if(dqTail < dqSize) Then
        if(Text0 = "_") Then
            dqEffect(0, dqTail) = eNone
            dqText(0, dqTail) = "_"
        Else
            dqEffect(0, dqTail) = Effect0
            dqText(0, dqTail) = ExpandLine(Text0, 0)
        End If

        if(Text1 = "_") Then
            dqEffect(1, dqTail) = eNone
            dqText(1, dqTail) = "_"
        Else
            dqEffect(1, dqTail) = Effect1
            dqText(1, dqTail) = ExpandLine(Text1, 1)
        End If

        if(Text2 = "_") Then
            dqEffect(2, dqTail) = eNone
            dqText(2, dqTail) = "_"
        Else
            dqEffect(2, dqTail) = Effect2
            dqText(2, dqTail) = Text2 'it is always 1 letter
        End If

        dqTimeOn(dqTail) = TimeOn
        dqbFlush(dqTail) = bFlush
        dqSound(dqTail) = Sound
        dqTail = dqTail + 1
        if(dqTail = 1) Then
            DMDHead()
        End If
    End If
End Sub

Sub DMDHead()
    Dim i
    deCount(0) = 0
    deCount(1) = 0
    deCount(2) = 0
    DMDEffectTimer.Interval = deSpeed

    For i = 0 to 2
        Select Case dqEffect(i, dqHead)
            Case eNone:deCountEnd(i) = 1
            Case eScrollLeft:deCountEnd(i) = Len(dqText(i, dqHead) )
            Case eScrollRight:deCountEnd(i) = Len(dqText(i, dqHead) )
            Case eBlink:deCountEnd(i) = int(dqTimeOn(dqHead) / deSpeed)
                deBlinkCycle(i) = 0
            Case eBlinkFast:deCountEnd(i) = int(dqTimeOn(dqHead) / deSpeed)
                deBlinkCycle(i) = 0
        End Select
    Next
    if(dqSound(dqHead) <> "") Then
        PlaySound(dqSound(dqHead) )
    End If
    DMDEffectTimer.Enabled = True
End Sub

Sub DMDEffectTimer_Timer()
    DMDEffectTimer.Enabled = False
    DMDProcessEffectOn()
End Sub

Sub DMDTimer_Timer()
    Dim Head
    DMDTimer.Enabled = False
    Head = dqHead
    dqHead = dqHead + 1
    if(dqHead = dqTail) Then
        if(dqbFlush(Head) = True) Then
            DMDFlush()
            DMDScore()
        Else
            dqHead = 0
            DMDHead()
        End If
    Else
        DMDHead()
    End If
End Sub

Sub DMDProcessEffectOn()
    Dim i
    Dim BlinkEffect
    Dim Temp

    BlinkEffect = False

    For i = 0 to 2
        if(deCount(i) <> deCountEnd(i) ) Then
            deCount(i) = deCount(i) + 1

            select case(dqEffect(i, dqHead) )
                case eNone:
                    Temp = dqText(i, dqHead)
                case eScrollLeft:
                    Temp = Right(dLine(i), dCharsPerLine(i) - 1)
                    Temp = Temp & Mid(dqText(i, dqHead), deCount(i), 1)
                case eScrollRight:
                    Temp = Mid(dqText(i, dqHead), (dCharsPerLine(i) + 1) - deCount(i), 1)
                    Temp = Temp & Left(dLine(i), dCharsPerLine(i) - 1)
                case eBlink:
                    BlinkEffect = True
                    if((deCount(i) MOD deBlinkSlowRate) = 0) Then
                        deBlinkCycle(i) = deBlinkCycle(i) xor 1
                    End If

                    if(deBlinkCycle(i) = 0) Then
                        Temp = dqText(i, dqHead)
                    Else
                        Temp = Space(dCharsPerLine(i) )
                    End If
                case eBlinkFast:
                    BlinkEffect = True
                    if((deCount(i) MOD deBlinkFastRate) = 0) Then
                        deBlinkCycle(i) = deBlinkCycle(i) xor 1
                    End If

                    if(deBlinkCycle(i) = 0) Then
                        Temp = dqText(i, dqHead)
                    Else
                        Temp = Space(dCharsPerLine(i) )
                    End If
            End Select

            if(dqText(i, dqHead) <> "_") Then
                dLine(i) = Temp
                DMDUpdate i
            End If
        End If
    Next

    if(deCount(0) = deCountEnd(0) ) and(deCount(1) = deCountEnd(1) ) and(deCount(2) = deCountEnd(2) ) Then

        if(dqTimeOn(dqHead) = 0) Then
            DMDFlush()
        Else
            if(BlinkEffect = True) Then
                DMDTimer.Interval = 10
            Else
                DMDTimer.Interval = dqTimeOn(dqHead)
            End If

            DMDTimer.Enabled = True
        End If
    Else
        DMDEffectTimer.Enabled = True
    End If
End Sub

Function ExpandLine(TempStr, id) 'id is the number of the dmd line
    If TempStr = "" Then
        TempStr = Space(dCharsPerLine(id) )
    Else
        if(Len(TempStr) > Space(dCharsPerLine(id) ) ) Then
            TempStr = Left(TempStr, Space(dCharsPerLine(id) ) )
        Else
            if(Len(TempStr) < dCharsPerLine(id) ) Then
                TempStr = TempStr & Space(dCharsPerLine(id) - Len(TempStr) )
            End If
        End If
    End If
    ExpandLine = TempStr
End Function

Function FormatScore(ByVal Num) 'it returns a string with commas (as in Black's original font)
    dim i
    dim NumString

    NumString = CStr(abs(Num) )

    For i = Len(NumString) -3 to 1 step -3
        if IsNumeric(mid(NumString, i, 1) ) then
            NumString = left(NumString, i-1) & chr(asc(mid(NumString, i, 1) ) + 48) & right(NumString, Len(NumString) - i)
        end if
    Next
    FormatScore = NumString
End function

Function CenterLine(id, NumString)
    Dim Temp, TempStr
    Temp = (dCharsPerLine(id) - Len(NumString) ) \ 2
    If(Temp + Temp + Len(NumString) ) < dCharsPerLine(id) Then
        TempStr = " " & Space(Temp) & NumString & Space(Temp)
    Else
        TempStr = Space(Temp) & NumString & Space(Temp)
    End If
    CenterLine = TempStr
End Function

Function RightLine(id, NumString)
    Dim Temp, TempStr
    Temp = dCharsPerLine(id) - Len(NumString)
    TempStr = Space(Temp) & NumString
    RightLine = TempStr
End Function

'*********************
' Update DMD - reels
'*********************

Dim DesktopMode:DesktopMode = MortalKombat.ShowDT

Dim Digits(2)

DMDReels_Init

Sub DMDReels_Init
    If DesktopMode Then
        'Desktop
        Digits(0) = Array(EMReel36, EMReel35, EMReel34, EMReel33, EMReel32, EMReel31, EMReel0, EMReel1, EMReel2, EMReel3, EMReel4, EMReel5, EMReel6, EMReel7, EMReel8, EMReel9)
        Digits(1) = Array(EMReel10, EMReel11, EMReel12, EMReel13, EMReel14, EMReel15, EMReel16, EMReel17, EMReel18, EMReel19, EMReel20, EMReel21, EMReel22, EMReel23, EMReel24, EMReel25, EMReel26, EMReel27, EMReel28, EMReel29)
        Digits(2) = Array(EMReel30)
        EMReel37.visible = 0
    Else
        'FS
        Digits(0) = Array(EMReel73, EMReel72, EMReel71, EMReel70, EMReel69, EMReel68, EMReel39, EMReel40, EMReel41, EMReel42, EMReel43, EMReel44, EMReel45, EMReel46, EMReel47, EMReel38)
        Digits(1) = Array(EMReel57, EMReel58, EMReel59, EMReel60, EMReel61, EMReel62, EMReel63, EMReel64, EMReel65, EMReel66, EMReel67, EMReel56, EMReel48, EMReel49, EMReel50, EMReel51, EMReel52, EMReel53, EMReel54, EMReel55)
        Digits(2) = Array(EMReel37)
        EMReel30.visible = 0
    End If
End Sub

Sub DMDUpdate(id)
    Dim digit, value
    If id < 2 Then 'text reels
        For digit = 0 to dCharsPerLine(id) -1
            value = ASC(mid(dLine(id), digit + 1, 1) ) -32
            Digits(id) (digit).SetValue value
        Next
    Else 'backdrop reel for animatons
        If dLine(2) = "" OR dLine(2) = " " Then
            value = 0
        Else
            value = dLine(2)
        End If
        Digits(2) (0).SetValue value
    End If
End Sub

'****************************************
' Real Time updatess using the GameTimer
'****************************************
'used for all the real time updates

Sub GameTimer_Timer
    RollingUpdate
    UpdateFlipperLogo
    Magnets_Ini
    GateL.RotX = -(Spinner3.currentangle)
    GateR.RotX = -(Spinner2.currentangle)
    KungLaoP.RotY = -(Spinner1.currentangle)
    Sombrero.RotY = -(Spinner1.currentangle)
    KungBase.RotY = -(Spinner1.currentangle)
  If BallsOnPlayfield > 0 Then 
     HologramAnim2.Enabled = 0
  End If
End Sub

'***************************************************
'     JP's VP10 Flashers v2 for original tables
'       Based on PD's Fading Light System
' SetFlash 0 is Off, 1 is On, 2 is blinking
'***************************************************

Dim FlashState(100), FadingLevel(100), FlashSpeedUp(100), FlashSpeedDown(100), FlashMin(100), FlashMax(100), FlashLevel(100), FlashRepeat(100)

InitFlashers()
 
Sub FlashTimer_Timer()
    Flash 1, Flasher1
    Flash 2, Flasher2
    Flash 3, MK1Light
    Flash 4, MK2Light
    Flash 5, MK3Light
    Flash 6, Flasher3
    Flash 7, Flasher4
    Flash 8, Flasher5
'    Flash 9, Flasher9
'    Flash 10, Flasher10
'    Flash 11, Flasher11
'    Flash 12, Flasher12
'    Flash 13, FlasherExitHole
End Sub

Sub InitFlashers()
    Dim x
    For x = 0 to 100
        FlashState(x) = 0        ' current state: 0 off, 1 on, 2 blinking
        FlashSpeedUp(x) = 0.5    ' faster speed when turning on the flasher
        FlashSpeedDown(x) = 0.25 ' slower speed when turning off the flasher
        FlashMax(x) = 1          ' the maximum value when on, usually 1
        FlashMin(x) = 0          ' the minimum value when off, usually 0
        FlashLevel(x) = 0        ' the intensity of the flashers, usually from 0 to 1
        FlashRepeat(x) = 20      ' the default number of blinks in the state 2
    Next
    ' Fast flashers - the dome flashers
    For x = 9 to 12
        FlashSpeedUp(x) = 1     'instant on
        FlashSpeedDown(x) = 0.5 ' only 2 step
    Next
    FlashTimer.Interval = 20
    FlashTimer.Enabled = True
End Sub

Sub AllFlashersOff
    Dim x
    For x = 0 to 100
        SetFlash x, 0
    Next
End Sub

Sub SetFlash(n, value)
	If value <> FlashState(n) Then
        FlashState(n) = value
        FadingLevel(n) = value
    End If
End Sub

Sub Flash(n, object)
    Select Case FadingLevel(n)
        Case 0 'off
            FlashLevel(n) = FlashLevel(n) - FlashSpeedDown(n)
            If FlashLevel(n) < FlashMin(n) Then
                FlashLevel(n) = FlashMin(n)
                FadingLevel(n) = -1 'stopped
            End if
            Object.IntensityScale = FlashLevel(n)
        Case 1 ' on
            FlashLevel(n) = FlashLevel(n) + FlashSpeedUp(n)
            If FlashLevel(n) > FlashMax(n) Then
                FlashLevel(n) = FlashMax(n)
                FadingLevel(n) = -1 'stopped
            End if
            Object.IntensityScale = FlashLevel(n)
        Case 2 'blink -turn on flasher
            FlashLevel(n) = FlashLevel(n) + FlashSpeedUp(n)
            If FlashLevel(n) > FlashMax(n) Then
                FlashLevel(n) = FlashMax(n)
                FadingLevel(n) = 3
            End if
            Object.IntensityScale = FlashLevel(n)
        Case 3 'blink -turn off flasher
            FlashLevel(n) = FlashLevel(n) - FlashSpeedDown(n)
            If FlashLevel(n) < FlashMin(n) Then
                FlashLevel(n) = FlashMin(n)
                FlashRepeat(n) = FlashRepeat(n) - 1
                If FlashRepeat(n) = 0 Then
                    FlashRepeat(n) = 20 'reset for next time
                    FlashState(n) = 0
                    FadingLevel(n) = -1
                Else
                    FadingLevel(n) = 2
                End If
            End if
            Object.IntensityScale = FlashLevel(n)
    End Select
End Sub

Sub Flashm(n, object) 'multiple flashers, it just sets the flashlevel
    Object.IntensityScale = FlashLevel(n)
End Sub

'******************************************
' Change light color - simulate color leds
' changes the light color and state
' colors: red, orange, yellow, green, blue, white
'******************************************

Sub SetLightColor(n, col, stat)
    Select Case col
        Case "red"
            n.color = RGB(18, 0, 0)
            n.colorfull = RGB(255, 0, 0)
        Case "orange"
            n.color = RGB(18, 3, 0)
            n.colorfull = RGB(255, 64, 0)
        Case "yellow"
            n.color = RGB(18, 18, 0)
            n.colorfull = RGB(255, 255, 0)
        Case "green"
            n.color = RGB(0, 18, 0)
            n.colorfull = RGB(0, 255, 0)
        Case "blue"
            n.color = RGB(0, 18, 18)
            n.colorfull = RGB(0, 255, 255)
        Case "white"
            n.color = RGB(255, 252, 224)
            n.colorfull = RGB(193, 91, 0)
        Case "purple"
            n.color = RGB(128, 0, 128)
            n.colorfull = RGB(255, 0, 255)
        Case "amber"
            n.color = RGB(193, 49, 0)
            n.colorfull = RGB(255, 153, 0)
    End Select
    If stat <> -1 Then
        n.State = 0
        n.State = stat
    End If
End Sub

'********************************************************************************************
' Only for VPX 10.2 and higher.
' FlashForMs will blink light or a flasher for TotalPeriod(ms) at rate of BlinkPeriod(ms)
' When TotalPeriod done, light or flasher will be set to FinalState value where
' Final State values are:   0=Off, 1=On, 2=Return to previous State
'********************************************************************************************

Sub FlashForMs(MyLight, TotalPeriod, BlinkPeriod, FinalState) 'thanks gtxjoe for the first myVersion

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

' ********************************
'   Table info & Attract Mode
' ********************************

Sub ShowTableInfo
    'info goes in a loop only stopped by the credits and the startkey
    If Score(1) Then
       DMD CenterLine(0, "LAST SCORE"), CenterLine(1, "PLAYER1 " &FormatScore(Score(1) ) ), 0, eNone, eNone, eNone, 3000, False, ""
       DMD_DisplaySceneEx "DMD1.png", "LAST SCORE", 15, 4, "Score "&Score(1) , -1, -1, UltraDMD_Animation_ScrollOnLeft, 3000, UltraDMD_Animation_ScrollOffLeft
       DMD_DisplaySceneEx "DMD1.png", "", 15, 4, "", -1, -1, UltraDMD_Animation_None, 500, UltraDMD_Animation_None
    End If
    If Score(2) Then
       DMD CenterLine(0, "LAST SCORE"), CenterLine(1, "PLAYER2 " &FormatScore(Score(2) ) ), 0, eNone, eNone, eNone, 3000, False, ""
       DMD_DisplaySceneEx "DMD1.png", "LAST SCORE", 15, 4, "Score "&Score(2) , -1, -1, UltraDMD_Animation_ScrollOnLeft, 3000, UltraDMD_Animation_ScrollOffLeft
       DMD_DisplaySceneEx "DMD1.png", "", 15, 4, "", -1, -1, UltraDMD_Animation_None, 500, UltraDMD_Animation_None
    End If
    If Score(3) Then
       DMD CenterLine(0, "LAST SCORE"), CenterLine(1, "PLAYER3 " &FormatScore(Score(3) ) ), 0, eNone, eNone, eNone, 3000, False, ""
       DMD_DisplaySceneEx "DMD1.png", "LAST SCORE", 15, 4, "Score "&Score(3) , -1, -1, UltraDMD_Animation_ScrollOnLeft, 3000, UltraDMD_Animation_ScrollOffLeft
       DMD_DisplaySceneEx "DMD1.png", "", 15, 4, "", -1, -1, UltraDMD_Animation_None, 500, UltraDMD_Animation_None
    End If
    If Score(4) Then
       DMD CenterLine(0, "LAST SCORE"), CenterLine(1, "PLAYER4 " &FormatScore(Score(4) ) ), 0, eNone, eNone, eNone, 3000, False, ""
       DMD_DisplaySceneEx "DMD1.png", "LAST SCORE", 15, 4, "Score "&Score(4) , -1, -1, UltraDMD_Animation_ScrollOnLeft, 3000, UltraDMD_Animation_ScrollOffLeft
       DMD_DisplaySceneEx "DMD1.png", "", 15, 4, "", -1, -1, UltraDMD_Animation_None, 500, UltraDMD_Animation_None
    End If
    DMD "", "", 3, eNone, eNone, eBlink, 2000, False, ""
    If bFreePlay Then
        DMD "", CenterLine(1, "FREE PLAY"), 0, eNone, eBlink, eNone, 2000, False, ""
    Else
        If Credits > 0 Then
            DMD CenterLine(0, "CREDITS " & Credits), CenterLine(1, "PRESS START"), 0, eNone, eBlink, eNone, 2000, False, ""
        Else
            DMD CenterLine(0, "CREDITS " & Credits), CenterLine(1, "INSERT COIN"), 0, eNone, eBlink, eNone, 2000, False, ""
        End If
    End If

     DMD_DisplaySceneEx "DMD1.png", "JAVIER", 15, 4, "PRESENT", -1, -1, UltraDMD_Animation_ScrollOnLeft, 5000, UltraDMD_Animation_ScrollOffRight
     DMD_DisplaySceneEx "DMD1.png", "", 15, 4, "", -1, -1, UltraDMD_Animation_None, 500, UltraDMD_Animation_None
     DMD_DisplaySceneEx "DMDMKII.png", "", 15, 4, "", -1, -1, UltraDMD_Animation_ScrollOnUp, 5000, UltraDMD_Animation_ScrollOffDown
     DMD_DisplaySceneEx "DMD1.png", "", 15, 4, "", -1, -1, UltraDMD_Animation_None, 100, UltraDMD_Animation_None  
     DMD_DisplaySceneEx "DMD1.png", "   Original Table   ", 15, 4, "    By Javier   ", -1, -1, UltraDMD_Animation_ScrollOnLeft, 3000, UltraDMD_Animation_ScrollOffRight
     DMD_DisplaySceneEx "DMD1.png", "", 15, 4, "", -1, -1, UltraDMD_Animation_None, 500, UltraDMD_Animation_None 


    DMD "", "", 3, eNone, eNone, eBlink, 2000, False, ""
    If Credits > 0 Then
        DMD CenterLine(0, "CREDITS " & Credits), CenterLine(1, "PRESS START"), 0, eNone, eBlink, eNone, 2000, False, ""
        DMD_DisplaySceneEx "DMD1.png", "CREDITS: "& Credits, 15, 4, "PRESS START", -1, -1, UltraDMD_Animation_None, 1000, UltraDMD_Animation_None
        DMD_DisplaySceneEx "DMD1.png", "PRESS START", 15, 4, "   BUTTON  ", -1, -1, UltraDMD_Animation_None, 700, UltraDMD_Animation_None
        DMD_DisplaySceneEx "DMD1.png", ":  ", 15, 4, "", -1, -1, UltraDMD_Animation_None, 300, UltraDMD_Animation_None
        DMD_DisplaySceneEx "DMD1.png", "PRESS START", 15, 4, "   BUTTON   ", -1, -1, UltraDMD_Animation_None, 700, UltraDMD_Animation_None
        DMD_DisplaySceneEx "DMD1.png", ":  ", 15, 4, "", -1, -1, UltraDMD_Animation_None, 300, UltraDMD_Animation_None
        DMD_DisplaySceneEx "DMD1.png", "PRESS START", 15, 4, "   BUTTON   ", -1, -1, UltraDMD_Animation_None, 700, UltraDMD_Animation_None
        DMD_DisplaySceneEx "DMD1.png", ":  ", 15, 4, "", -1, -1, UltraDMD_Animation_None, 300, UltraDMD_Animation_None
        DMD_DisplaySceneEx "DMD1.png", "PRESS START", 15, 4, "   BUTTON   ", -1, -1, UltraDMD_Animation_None, 700, UltraDMD_Animation_None
    Else
        DMD CenterLine(0, "CREDITS " & Credits), CenterLine(1, "INSERT COIN"), 0, eNone, eBlink, eNone, 2000, False, ""
        DMD_DisplaySceneEx "DMD1.png", "CREDITS "& Credits, 15, 4, "INSERT COIN", -1, -1, UltraDMD_Animation_None, 700, UltraDMD_Animation_None
        DMD_DisplaySceneEx "DMD1.png", "CREDITS "& Credits, 15, 4, " ", -1, -1, UltraDMD_Animation_None, 50, UltraDMD_Animation_None
        DMD_DisplaySceneEx "DMD1.png", "CREDITS "& Credits, 15, 4, "INSERT COIN", -1, -1, UltraDMD_Animation_None, 700, UltraDMD_Animation_None
        DMD_DisplaySceneEx "DMD1.png", "CREDITS "& Credits, 15, 4, " ", -1, -1, UltraDMD_Animation_None, 50, UltraDMD_Animation_None
        DMD_DisplaySceneEx "DMD1.png", "CREDITS "& Credits, 15, 4, "INSERT COIN", -1, -1, UltraDMD_Animation_None, 700, UltraDMD_Animation_None
        DMD_DisplaySceneEx "DMD1.png", "CREDITS "& Credits, 15, 4, " ", -1, -1, UltraDMD_Animation_None, 50, UltraDMD_Animation_None
        DMD_DisplaySceneEx "DMD1.png", "CREDITS "& Credits, 15, 4, "INSERT COIN", -1, -1, UltraDMD_Animation_None, 700, UltraDMD_Animation_None 
    End If


    DMD_DisplaySceneEx "DMD1.png", "HIGHSCORES", 15, 4, "", -1, -1, UltraDMD_Animation_ScrollOnLeft, 800, UltraDMD_Animation_None  
    DMD_DisplaySceneEx "DMD1.png", "", 15, 4, "", -1, -1, UltraDMD_Animation_None, 5, UltraDMD_Animation_None
    DMD_DisplaySceneEx "DMD1.png", "HIGHSCORES", 15, 4, "", -1, -1, UltraDMD_Animation_None, 100, UltraDMD_Animation_None  
    DMD_DisplaySceneEx "DMD1.png", "", 15, 4, "", -1, -1, UltraDMD_Animation_None, 5, UltraDMD_Animation_None
    DMD_DisplaySceneEx "DMD1.png", "HIGHSCORES", 15, 4, "", -1, -1, UltraDMD_Animation_None, 100, UltraDMD_Animation_None  
    DMD_DisplaySceneEx "DMD1.png", "", 15, 4, "", -1, -1, UltraDMD_Animation_None, 5, UltraDMD_Animation_None
    DMD_DisplaySceneEx "DMD1.png", "HIGHSCORES", 15, 4, "", -1, -1, UltraDMD_Animation_None, 100, UltraDMD_Animation_None  
    DMD_DisplaySceneEx "DMD1.png", "", 15, 4, "", -1, -1, UltraDMD_Animation_None, 5, UltraDMD_Animation_None
    DMD_DisplaySceneEx "DMD1.png", "HIGHSCORES", 15, 4, "", -1, -1, UltraDMD_Animation_None, 100, UltraDMD_Animation_None  
    DMD_DisplaySceneEx "DMD1.png", "", 15, 4, "", -1, -1, UltraDMD_Animation_None, 5, UltraDMD_Animation_None
    DMD_DisplaySceneEx "DMD1.png", "HIGHSCORES", 15, 4, "", -1, -1, UltraDMD_Animation_None, 100, UltraDMD_Animation_None  
    DMD_DisplaySceneEx "DMD1.png", "", 15, 4, "", -1, -1, UltraDMD_Animation_None, 5, UltraDMD_Animation_None
    DMD_DisplaySceneEx "DMD1.png", "HIGHSCORES", 15, 4, "", -1, -1, UltraDMD_Animation_None, 100, UltraDMD_Animation_None  
    DMD_DisplaySceneEx "DMD1.png", "", 15, 4, "", -1, -1, UltraDMD_Animation_None, 5, UltraDMD_Animation_None
    DMD_DisplaySceneEx "DMD1.png", "HIGHSCORES", 15, 4, "1>"&HighScoreName(0) & " " &HighScore(0) , -2, -1, UltraDMD_Animation_ScrollOnLeft, 3000, UltraDMD_Animation_ScrollOffLeft
    DMD_DisplaySceneEx "DMD1.png", "HIGHSCORES", 15, 4, "2>"&HighScoreName(1) & " " &HighScore(1) , -2, -1, UltraDMD_Animation_ScrollOnLeft, 3000, UltraDMD_Animation_ScrollOffLeft
    DMD_DisplaySceneEx "DMD1.png", "HIGHSCORES", 15, 4, "3>"&HighScoreName(2) & " " &HighScore(2) , -2, -1, UltraDMD_Animation_ScrollOnLeft, 3000, UltraDMD_Animation_ScrollOffLeft
    DMD_DisplaySceneEx "DMD1.png", "HIGHSCORES", 15, 4, "4>"&HighScoreName(3) & " " &HighScore(3) , -2, -1, UltraDMD_Animation_ScrollOnLeft, 3000, UltraDMD_Animation_ScrollOffLeft
    DMD_DisplaySceneEx "DMD1.png", "", 15, 4, "", -1, -1, UltraDMD_Animation_None, 3000, UltraDMD_Animation_None
    UltraDMD.DisplayScoreboard PlayersPlayingGame, CurrentPlayer, Score(1), Score(2), Score(3), Score(4), "Ball " & Balls, "Credits: " & Credits
    AttractAnimUltraDND.enabled = 1

    DMD "", "", 1, eNone, eNone, eNone, 3000, False, ""
    DMD "", "", 2, eNone, eNone, eNone, 4000, False, ""
    DMD CenterLine(0, "HIGHSCORES"), Space(dCharsPerLine(1) ), 0, eScrollLeft, eScrollLeft, eNone, 20, False, ""
    DMD CenterLine(0, "HIGHSCORES"), "", 0, eBlinkFast, eNone, eNone, 1000, False, ""
    DMD CenterLine(0, "HIGHSCORES"), "1> " &HighScoreName(0) & " " &FormatScore(HighScore(0) ), 0, eNone, eScrollLeft, eNone, 2000, False, ""

    DMD "_", "2> " &HighScoreName(1) & " " &FormatScore(HighScore(1) ), 0, eNone, eScrollLeft, eNone, 2000, False, ""
    DMD "_", "3> " &HighScoreName(2) & " " &FormatScore(HighScore(2) ), 0, eNone, eScrollLeft, eNone, 2000, False, ""
    DMD "_", "4> " &HighScoreName(3) & " " &FormatScore(HighScore(3) ), 0, eNone, eScrollLeft, eNone, 2000, False, ""
    DMD Space(dCharsPerLine(0) ), Space(dCharsPerLine(1) ), 0, eScrollLeft, eScrollLeft, eNone, 500, False, ""
End Sub

Sub StartAttractMode(dummy)
    ChangeSong
    StartLightSeq
    DMDFlush
    ShowTableInfo
End Sub

Sub StopAttractMode
    Dim bulb
    DMDFlush
    LightSeqAttract.StopPlay
'StopSong
End Sub

Sub StartLightSeq()
    'lights sequences
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

'***********************************************************************
' *********************************************************************
'                     Table Specific Script Starts Here
' *********************************************************************
'***********************************************************************

' tables walls and animations
Sub VPObjects_Init
End Sub

' tables variables and modes init

Dim Coins(2)
Dim BumperBonus, bRampBonus, CharacterBonus, TargetBonus
Dim BallInHole, bLockEnabled, LockedBalls

Sub Game_Init()
    StopAnimUDMD
    bExtraBallWonThisBall = False
    TurnOffPlayfieldLights()
    'Play some Music
    ChangeSong
    'Init Variables
    Jackpot = 250000
    bumperHits = 100
    ExtraBallCount = 0
    BumperBonus = 0
    BallInHole = 0
    CharacterBonus = 0
    TargetBonus = 0
    bLockEnabled = False
    LockedBalls = 0
    bRampBonus = FALSE
'    bLoopBonus = FALSE
    LightSkillShot.state  = 1
    SwSkillShot.enabled = 1
    HologramAnim2.Enabled = 0
    HologramRampF.Visible = 0
     l1.state = 2 
     l2.state = 2  
     l3.state = 2
     l4.state = 2 
     l5.state = 2  
     l6.state = 2
     l7.state = 2 
     l8.state = 2  
     l9.state = 2
     LockLight.state = 0
     LockLight1.state = 2
     LockLight2.state = 2
     LockLight3.state = 2
     MultiLight1.state = 0
     MultiLight2.state = 0
     MultiLight3.state = 0
     SetFlash 3, 0
     SetFlash 4, 0
     SetFlash 5, 0
     If Sw25.Isdropped = 1 Then DOF 114, DOFPulse
     Sw25.Isdropped = 0
     Post.Isdropped= 0
     ScorpionD = 0
     RaidenD = 0
     KungLaoD = 0
     JaxD = 0
     MillenaD = 0
     KitanaD = 0
     SmokeD = 0
     KintaroD = 0
     ShaoKahnD = 0
     SubZeroD = 0
     JohnnyCajeD = 0
     ReptileD = 0
     BarakaD = 0
     ShangTsungD = 0
     JadeD = 0
     MKIC = 0
     MKIIC = 0
     MKIIIC = 0
    AttractAnimUltraDND.enabled = 0
     CharactersResetLight()
End Sub


Sub ResetSkillShotTimer_Timer
End Sub

Sub TurnOffPlayfieldLights()
    Dim a
    For each a in aLights
        a.State = 0
    Next
    SetFlash 1, 0
    SetFlash 2, 0
    SetFlash 3, 0
    SetFlash 4, 0
    SetFlash 5, 0
End Sub

Sub ResetNewBallLights()
'	LightArrow1.State = 2
'	LightArrow6.State = 2
'	l53.State = 2
End Sub

' *********************************************************************
'                        Table Object Hit Events
'
' Any target hit Sub will follow this:
' Any target hit Sub will follow this:
' - play a sound
' - do some physical movement
' - add a score, bonus
' - check some variables/modes this trigger is a member of
' - set the "LastSwicthHit" variable in case it is needed later
' *********************************************************************

' Slingshots has been hit

Dim LStep, RStep

Sub LeftSlingShot_Slingshot
    If Tilted Then Exit Sub
    PlaySound SoundFXDOF("fx_slingshot",103,DOFPulse,DOFContactors), 0, 1, -0.05, 0.05
	DOF 105, DOFPulse
    LeftSling4.Visible = 1:LeftSling1.Visible = 0
    Lemk.RotX = 26
    LStep = 0
    LeftSlingShot.TimerEnabled = True
    ' add some points
    AddScore 110
    ' add some effect to the table?
    Gi2.State = 0
    ' remember last trigger hit by the ball
    LastSwitchHit = "LeftSlingShot"
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 1:LeftSLing4.Visible = 0:LeftSLing3.Visible = 1:Lemk.RotX = 14
        Case 2:LeftSLing3.Visible = 0:LeftSLing2.Visible = 1:Lemk.RotX = 2
        Case 3:LeftSLing2.Visible = 0:LeftSling1.Visible = 1:Lemk.RotX = -10:Gi2.State = 1:LeftSlingShot.TimerEnabled = False
    End Select
    LStep = LStep + 1
End Sub

Sub RightSlingShot_Slingshot
    If Tilted Then Exit Sub
    PlaySound SoundFXDOF("fx_slingshot",104,DOFPulse,DOFContactors), 0, 1, 0.05, 0.05
	DOF 106, DOFPulse
    RightSling4.Visible = 1:RightSling1.Visible = 0
    Remk.RotX = 26
    RStep = 0
    RightSlingShot.TimerEnabled = True
    ' add some points
    AddScore 110
    ' add some effect to the table?
    Gi1.State = 0
    ' remember last trigger hit by the ball
    LastSwitchHit = "RightSlingShot"
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 1:RightSLing4.Visible = 0:RightSLing3.Visible = 1:Remk.RotX = 14
        Case 2:RightSLing3.Visible = 0:RightSLing2.Visible = 1:Remk.RotX = 2
        Case 3:RightSLing2.Visible = 0:RightSLing1.Visible = 1:Remk.RotX = -10:Gi1.State = 1:RightSlingShot.TimerEnabled = False
    End Select
    RStep = RStep + 1
End Sub

'**********
' Spinner
'**********

Sub Spinner_Spin()
    If Tilted Then Exit Sub
	PlaySound"fx_spinner", 0, 1, 0.1
    AddScore 1000
End Sub

'********
' Bumper
'********

Dim bumperHits
 
Sub Bumper1b_Hit
    If Tilted Then Exit Sub
    FlasherB.State = 1
    Bumper1bL.state = 1
    vpmTimer.AddTimer 50, "ResetBumpersLights"
    PlaySound SoundFXDOF("fx_bumper",107,DOFPulse,DOFContactors), 0, 1, pan(ActiveBall)
	DOF 110, DOFPulse
    AddScore 500 + 4500 * Flashstate(1) 'a bumper scores 100 points and 1000 points when lit
    bumperHits = bumperHits - 1
    DMDFlush
    DMD RightLine(0, FormatScore(Score(Currentplayer) ) ), RightLine(1, bumperHits& " BUMPER HITS LEFT"), 0, eNone, eNone, eNone, 300, True, ""
    DMD_DisplaySceneEx "DMD1.png", "BUMPER LEFT" , 15, 4, "HITS "&bumperHits , -1, -1, UltraDMD_Animation_None, 100, UltraDMD_Animation_None
    CheckBumpers
End Sub

Sub Bumper2b_Hit
    If Tilted Then Exit Sub
    FlasherB.State = 1
    Bumper2bL.state = 1
    vpmTimer.AddTimer 50, "ResetBumpersLights"
    PlaySound SoundFXDOF("fx_bumper",109,DOFPulse,DOFContactors), 0, 1, pan(ActiveBall)
	DOF 112, DOFPulse
    AddScore 500 + 4500 * Flashstate(2) 'a bumper scores 100 points and 1000 points when lit
    bumperHits = bumperHits - 1
    DMDFlush
    DMD RightLine(0, FormatScore(Score(Currentplayer) ) ), RightLine(1, bumperHits& " BUMPER HITS RIGHT"), 0, eNone, eNone, eNone, 300, True, ""
    DMD_DisplaySceneEx "DMD1.png", "BUMPER RIGHT" , 15, 4, "HITS "&bumperHits  , -1, -1, UltraDMD_Animation_None, 100, UltraDMD_Animation_None
    CheckBumpers
End Sub

Sub Bumper3b_Hit
    If Tilted Then Exit Sub
    FlasherB.State = 1
    Bumper3bL.state = 1
    vpmTimer.AddTimer 50, "ResetBumpersLights"
    PlaySound SoundFXDOF("fx_bumper",108,DOFPulse,DOFContactors), 0, 1, pan(ActiveBall)
	DOF 111, DOFPulse
    AddScore 500 + 4500 * Flashstate(3) 'a bumper scores 100 points and 1000 points when lit
    bumperHits = bumperHits - 1
    DMDFlush
    DMD RightLine(0, FormatScore(Score(Currentplayer) ) ), RightLine(1, bumperHits& "BUMPER HITS BOTTOM"), 0, eNone, eNone, eNone, 300, True, ""
    DMD_DisplaySceneEx "DMD1.png", "BUMPER BOTTOM" , 15, 4, "HITS "&bumperHits  , -1, -1, UltraDMD_Animation_None, 100, UltraDMD_Animation_None
    CheckBumpers
End Sub

' Bumper Bonus
' 100000 i bonus after each 100 hits

Sub CheckBumpers()
 '   FlasherB.state=0
    If bumperHits <= 0 Then
        BumperBonus = BumperBonus + 1
        DMD "_", RightLine(1, "BUMPERS BONUS " & BumperBonus), 0, eNone, eBlink, eNone, 500, True, ""
        bumperHits = 100
    ' do something more
    End If
End Sub

Sub ResetBumpers()
    bumperHits = 100
'    FlasherB.state = 0
'    SetFlash 6, 0
'    SetFlash 7, 0
'    SetFlash 8, 0
End Sub

Sub ResetBumpersLights(dummy)
    FlasherB.State = 0
    Bumper1bL.state = 0 
    Bumper2bL.state = 0
    Bumper3bL.state = 0
End Sub

'***********
'Skillshot
'***********
Sub SwSkillShot_hit()
   vpmTimer.AddTimer 500, "SwSkillShotUpdate"
   AddScore 1000
   DMDFlush
   DMD RightLine(0, ""), CenterLine(1, ""), 5, eNone, eNone, eBlink, 1000, True, ""
   LightSkillShot.BlinkInterval = 150
   LightSkillShot.State = 2
   SwSkillShotLightTimer.enabled = 1
End Sub  

Sub SwSkillShotUpdate(dummy)
	DOF 130, DOFPulse
    SwSkillShot.kick 0, 55
    LaserKickP.TransY = 90 
    Playsound "bumper_retro"
End Sub


Sub SwSkillShot_Unhit()
    LaserKickP.TransY = 0
    PlaySound "ScorpionGetOverHere"
    LightSkillShot.State = 0  
End Sub


sub SwSkillShotLightTimer_timer()
    LightSkillShot.State = 0
    me.enabled=0
end sub

'*********
'Targgets
'*********
Dim bsLocksTargets1, bsLocksTargets2, bsLocksTargets3
bsLocksTargets1 = 0: bsLocksTargets2 = 0: bsLocksTargets3 = 0
Sub Sw1_hit
    If Tilted Then Exit Sub
    PlaySound SoundFXDOF("fx_rubber",117,DOFPulse,DOFTargets)
    AddScore 1000  
  If bMultiBallMode Then Exit Sub
    l1.state = 1 
    DMD RightLine(0, FormatScore(Score(Currentplayer) ) ), CenterLine(1, " M "), 0, eNone, eBlink, eNone, 500, True, "M" 
    MKIUpdate() 
    TargetBonus = TargetBonus + 1
End Sub

Sub Sw2_hit
    If Tilted Then Exit Sub
    PlaySound SoundFXDOF("fx_rubber",117,DOFPulse,DOFTargets)
    AddScore 1000 
  If bMultiBallMode Then Exit Sub 
    l2.state = 1 
    DMD RightLine(0, FormatScore(Score(Currentplayer) ) ), CenterLine(1, " K "), 0, eNone, eBlink, eNone, 500, True, "K" 
    MKIUpdate() 
    TargetBonus = TargetBonus + 1
End Sub

Sub Sw3_hit
    If Tilted Then Exit Sub
    PlaySound SoundFXDOF("fx_rubber",117,DOFPulse,DOFTargets)
    AddScore 1000  
  If bMultiBallMode Then Exit Sub
    l3.state = 1 
    DMD RightLine(0, FormatScore(Score(Currentplayer) ) ), CenterLine(1, " I "), 0, eNone, eBlink, eNone, 500, True, "One" 
    MKIUpdate() 
    TargetBonus = TargetBonus + 1
End Sub

Sub MKIUpdate
  If bMultiBallMode Then Exit Sub
  If (l1.state + l2.state + l3.state) = 3 Then
     AddScore 10000
     DMDFlush
     DMD RightLine(0, ""), CenterLine(1, ""), 25, eNone, eBlink, eNone, 1000, True, "Excellent"
     bsLocksTargets1 = 1
     SetFlash 3, 2
     MKIReset()
     BackFlashEffect 1
  End If
End Sub

Dim MKIC: MKIC = 0
Sub MKIReset()
  if bsLocksTargets1 = 1 Then
     l1.state = 2
     l2.state = 2
     l3.state = 2
     SetFlash 3, 1
     MKIC = 1
     MortalKombatTargets
  End If
End Sub

Sub Sw4_hit
    If Tilted Then Exit Sub
    PlaySound SoundFXDOF("fx_rubber",113,DOFPulse,DOFTargets)
    AddScore 1000 
  If bMultiBallMode Then Exit Sub 
    l4.state = 1 
    DMD RightLine(0, FormatScore(Score(Currentplayer) ) ), CenterLine(1, " III "), 0, eNone, eBlink, eNone, 500, True, "Three" 
    MKIIIUpdate() 
    TargetBonus = TargetBonus + 1
End Sub

Sub Sw5_hit
    If Tilted Then Exit Sub
    PlaySound SoundFXDOF("fx_rubber",113,DOFPulse,DOFTargets)
    AddScore 1000 
  If bMultiBallMode Then Exit Sub 
    l5.state = 1 
    DMD RightLine(0, FormatScore(Score(Currentplayer) ) ), CenterLine(1, " K "), 0, eNone, eBlink, eNone, 500, True, "K" 
    MKIIIUpdate() 
    TargetBonus = TargetBonus + 1
End Sub

Sub Sw6_hit
    If Tilted Then Exit Sub
    PlaySound SoundFXDOF("fx_rubber",113,DOFPulse,DOFTargets)
    AddScore 1000  
  If bMultiBallMode Then Exit Sub
    l6.state = 1 
    DMD RightLine(0, FormatScore(Score(Currentplayer) ) ), CenterLine(1, " M "), 0, eNone, eBlink, eNone, 500, True, "M" 
    MKIIIUpdate() 
    TargetBonus = TargetBonus + 1
End Sub

Sub MKIIIUpdate()
  If bMultiBallMode Then Exit Sub
  If (l4.state + l5.state + l6.state) = 3 Then
     AddScore 10000
     DMDFlush
     DMD RightLine(0, ""), CenterLine(1, ""), 27, eNone, eBlink, eNone, 1000, True, "Excellent"
     bsLocksTargets3 = 1
     SetFlash 5, 2
     MKIIIReset
     BackFlashEffect 1
  End If
End Sub

Dim MKIIIC: MKIIIC = 0
Sub MKIIIReset
  if bsLocksTargets3 = 1 Then
     l4.state = 2
     l5.state = 2
     l6.state = 2
     MKIIIC = 1
     SetFlash 5, 1
     MortalKombatTargets
  End If
End Sub

Sub Sw7_hit
    If Tilted Then Exit Sub
    PlaySound SoundFXDOF("fx_rubber",114,DOFPulse,DOFTargets)
    AddScore 1000 
  If bMultiBallMode Then Exit Sub 
    l7.state = 1 
    DMD RightLine(0, FormatScore(Score(Currentplayer) ) ), CenterLine(1, " II "), 0, eNone, eBlink, eNone, 500, True, "Two" 
    MKIIUpdate() 
    TargetBonus = TargetBonus + 1
End Sub

Sub Sw8_hit
    If Tilted Then Exit Sub
    PlaySound SoundFXDOF("fx_rubber",114,DOFPulse,DOFTargets)
    AddScore 1000  
  If bMultiBallMode Then Exit Sub
    l8.state = 1 
    DMD RightLine(0, FormatScore(Score(Currentplayer) ) ), CenterLine(1, " K "), 0, eNone, eBlink, eNone, 500, True, "K" 
    MKIIUpdate() 
    TargetBonus = TargetBonus + 1
End Sub

Sub Sw9_hit
    If Tilted Then Exit Sub
    PlaySound SoundFXDOF("fx_rubber",114,DOFPulse,DOFTargets)
    AddScore 1000  
  If bMultiBallMode Then Exit Sub
    l9.state = 1 
    DMD RightLine(0, FormatScore(Score(Currentplayer) ) ), CenterLine(1, " M "), 0, eNone, eBlink, eNone, 500, True, "M" 
    MKIIUpdate() 
    TargetBonus = TargetBonus + 1
End Sub

Sub MKIIUpdate
  If bMultiBallMode Then Exit Sub
  If (l7.state + l8.state + l9.state) = 3 Then
     AddScore 10000
     DMDFlush
     DMD RightLine(0, ""), CenterLine(1, ""), 26, eNone, eBlink, eNone, 1000, True, "Excellent"
     bsLocksTargets2 = 1
     SetFlash 4, 2
     MKIIReset
     BackFlashEffect 1
  End If 
End Sub

Dim MKIIC: MKIIC = 0
Sub MKIIReset
  if bsLocksTargets2 = 1 Then
     l7.state = 2
     l8.state = 2
     l9.state = 2
     MKIIC = 1
     SetFlash 4, 1
     MortalKombatTargets
  End If
End Sub

Sub Sw10_hit
   AddScore 2000
  If LightSkillShot.state = 1 Then 
    DMDFlush
    DMD RightLine(0, FormatScore(Currentplayer) ), CenterLine(1, " FIGHT "), 0, eNone, eNone, eNone, 600, True, "Fight"
 End If
  If LightSkillShot.state = 0 Then  
    DMDFlush
    DMD RightLine(0, ""), CenterLine(1, ""), 24, eNone, eBlink, eNone, 1000, True, "Scorpion"
    LightSkillShot.state = 2
    SkillshotUpdate()
 End If
End Sub

Sub SkillshotUpdate()
    LightSkillShot.state = 1
End Sub

Sub Sw11_hit()
  vpmTimer.AddTimer 500, "Sw11Update"
  Playsound "Fight"
End Sub

Sub Sw11Update(dummy)
    Sw11.kick 0, 48, 1.5
	DOF 128, DOFPulse
	DOF 125, DOFPulse
    Playsound "fx_rampR"
    AddScore 2000
   If bMultiBallMode = True and LJ5.state = 2 Then
      LJ5.state = 1
      AwardJackpot()
      ChectSuperJackpot()
   End If 
End Sub

Sub Sw11_Unhit()
  Playsound "ScorpionWins"
  SetFlash 6, 2
  DOF 208, DOFPulse
  LightEffect 1
End Sub

Sub Sw12_Hit()
    If Tilted Then Exit Sub
    PlaySound "fx_hole-enter", 0, 1, -0.1, 0.25
    AddScore 2000
    Me.DestroyBall
    SetFlash 8, 2
	DOF 209, DOFPulse
  If bMultiBallMode = False Then
 '    startB2S(20)
     StartSlotmachine
     l18.state = 0
	 If Sw25.Isdropped = 1 Then DOF 114, DOFPulse
     Sw25.Isdropped = 0
     CharacterBonus = CharacterBonus + 1
  Else
     vpmtimer.addtimer 1000, "Sw13Hole '"
     AddScore 2000
   if bMultiBallMode And LJ3.state = 2 Then
     vpmtimer.addtimer 500, "SetFlash 1, 2 '"
     vpmtimer.addtimer 500, "SetFlash 2, 2 '"
     LJ3.state = 1
     AwardJackpot
   if bMultiBallMode And LJ3.state = 1 Then
      ChectSuperJackpot()
    End If
    End If
   End If
    LastSwitchHit = "Sw12"
End Sub

Sub Sw13_Hit()
    EnableBallSaver 2
    PlaySound "fx_hole-enter", 0, 1, 0.1
    Me.Destroyball
    AddScore 3000
    If Not Tilted Then
        If bLockEnabled Then
            LockedBalls = LockedBalls + 1
            CheckLockMulltiball()
            DMD "_", CenterLine(1, "BALL " & LockedBalls & " IS LOCKED"), "", eNone, eNone, eNone, 1600, True, ""
            DMD_DisplaySceneEx "DMD1.png", "BALL  " & LockedBalls , 15, 4, "IS LOCKED", -1, -1, UltraDMD_Animation_None, 2000, UltraDMD_Animation_None
            If LockedBalls = 3 Then
                MKMultiball
            End IF
        End If
   ' End If
   End If
'    vpmtimer.addtimer 2500, "Sw13Hole '"
    ' remember last trigger hit by the ball
    LastSwitchHit = "Sw13"
End Sub

Sub Sw13Hole()
	DOF 129, DOFPulse
	DOF 125, DOFPulse
    Sw14.createball 
    Sw14.kick 0, 38, 1.56
    Playsound "fx_vukout_LAH"
    BackFlashEffect 1
    KungLaoTrompo()
  If LockedBalls = 1 Then
     MultiLight1.state = 1
  End If

  If LockedBalls = 2 Then
     MultiLight2.state = 1
     MultiLight1.state = 1
  End If

  If LockedBalls = 3 Then
     MultiLight3.state = 1
     MultiLight2.state = 1
     MultiLight1.state = 1
  End If
End Sub

Sub CheckLockMulltiball()
  If LockedBalls = 3 Then
     MKMultiball
     vpmtimer.addtimer 2500, "Sw13Hole '" 
   Else
     Post.Isdropped = 0
     LockLight.state = 0
     LockLight1.state = 2
     LockLight2.state = 2
     LockLight3.state = 2
     bLockEnabled = False
     vpmtimer.addtimer 2500, "Sw13Hole '"
  End If   
End Sub  

'*********************
' Inlanes - Outlanes
'*********************

Sub Sw15_Hit()
	DOF 118, DOFPulse
    PlaySound "fx_sensor", 0, 1, pan(ActiveBall)
    If Tilted Then Exit Sub
    AddScore 3000
    l11.State = 1
    LastSwitchHit = "Sw15"
    CheckMultiplierUp
End Sub

Sub Sw16_Hit()
	DOF 119, DOFPulse
    PlaySound "fx_sensor", 0, 1, pan(ActiveBall)
    If Tilted Then Exit Sub
    AddScore 3000
    l12.State = 1
    CheckMultiplierUp
    LastSwitchHit = "Sw16"
End Sub

Sub Sw17_Hit()
	DOF 120, DOFPulse
    PlaySound "fx_sensor", 0, 1, pan(ActiveBall)
    If Tilted Then Exit Sub
    AddScore 10000
    SetFlash 9, 2:SetFlash 10, 2
    l13.State = 1
    CheckMultiplierUp
    LastSwitchHit = "Sw17"
End Sub

Sub Sw18_Hit()
'    startB2S(17)
    PlaySound "fx_sensor", 0, 1, pan(ActiveBall)
    If Tilted Then Exit Sub
    AddScore 3000
    l14.State = 1
    vpmTimer.AddTimer 100, "ComboblinkLight"
    LastSwitchHit = "Sw18"
    CheckMultiplier
End Sub

Sub Sw19_Hit()
'    startB2S(18)
    PlaySound "fx_sensor", 0, 1, pan(ActiveBall)
    If Tilted Then Exit Sub
    AddScore 3000
    l15.State = 1
    vpmTimer.AddTimer 100, "ComboblinkLight"
    CheckMultiplier
    LastSwitchHit = "Sw19"
End Sub

Sub Sw20_Hit()
    If Tilted Then Exit Sub
    AddScore 10000
    SetFlash 1, 2:SetFlash 2, 2
	DOF 206, DOFPulse
	DOF 207, DOFPulse
    l10.State = 1

    If LightSkillShot.state = 0 Then 
    PlaySound "Kahn_2", 0, 1, pan(ActiveBall)
    End If
    If LightSkillShot.state = 1 Then 
    PlaySound "scorp_spear", 0, 1, pan(ActiveBall)
    End If

    If LightSkillShot.State = 1  Or LightSkillShot.state = 2 Then: SwSkillShot.enabled = 1 end If
    If LightSkillShot.State = 0 Then: SwSkillShot.enabled = 0 end If

    LastSwitchHit = "Sw20"
    CheckMultiplier
End Sub

Sub Sw21a_Hit()
    LightEffect 3
'    startB2S(17)
 '   startB2S(18)
end Sub

Sub Sw21_Hit()
    PlaySound "Kahn_2", 0, 1, pan(ActiveBall)
    If Tilted Then Exit Sub
    AddScore 10000
    SetFlash 1, 2:SetFlash 2, 2
	DOF 206, DOFPulse
	DOF 207, DOFPulse
    LightEffect 3
    l17.State = 1
 '   startB2S(17)
 '   startB2S(18)
    CheckMultiplier
    LastSwitchHit = "Sw21"
End Sub

Sub Sw22_hit
    PlaySound "fx_sensor", 0, 1, pan(ActiveBall)
    If Tilted Then Exit Sub
   If bMultiBallMode and LJ2.state = 2 Then
      LJ2.state = 1
      AwardJackpot()
   If bMultiBallMode and LJ2.state = 1 Then
      ChectSuperJackpot()
    End If
   End If
End sub

Sub CheckMultiplier
    If(l10.State = 1) And(l14.State = 1) And(l15.State = 1) And(l17.State = 1) Then
        AddScore 50000
        DMDFlush
        DMD RightLine(0, ""), CenterLine(1, ""), 32, eNone, eBlink, eNone, 800, True, "Outstanding"
        LightSeqLanes.Play SeqRandom, 5, , 2000
        IncrementBonusMultiplier()
        l10.State = 0
        l14.State = 0
        l15.State = 0
        l17.State = 0
    End If
End Sub

Sub CheckMultiplierUp
    If(l11.State = 1) And(l12.State = 1) And(l13.State = 1) Then
        DMDFlush
        DMD RightLine(0, ""), CenterLine(1, ""), 32, eNone, eBlink, eNone, 800, True, "Outstanding"
        AddScore 50000
        LightSeqLanesUP.Play SeqRandom, 5, , 2000
        IncrementBonusMultiplier()
        l11.State = 0
        l12.State = 0
        l13.State = 0
    End If
End Sub

Sub Sw25_hit()
    If Tilted Then Exit Sub
    PlaySound SoundFXDOF("fx_target",114,DOFPulse,DOFTargets)
    AddScore 2000
    DMDFlush
    DMD RightLine(0, FormatScore(Score(Currentplayer) ) ), CenterLine(1, " RANDOM IS READY "), 0, eNone, eNone, eNone, 800, True, ""   
    l18.state = 2
End Sub 

Dim ExtraBallCount
'Left Ramp  
Sub sw26_Hit
    PlaySound "fx_rampL", 0, 1, pan(ActiveBall)
    If Tilted Then Exit Sub
    AddScore 2000
    SetFlash 2, 2
	DOF 207, DOFPulse
'    startB2S(17)
    vpmTimer.AddTimer 100, "ComboblinkLight"
'    LightDragon.BlinkInterval = 160
    ExtraBallCount = ExtraBallCount + 1

If bMultiBallMode Then
 If bSuperJackpotMode = false And LJ1.state = 2 Then
     LJ1.state = 1
     AwardJackpot
     ChectSuperJackpot()
   Else
 If LJ1.state = 1 Then
     ChectSuperJackpot()
  End If
 End If
End If

   If ExtraBallCount = 8 or ExtraBallCount = 16 Or ExtraBallCount = 24 Then
      AwardExtraBall()
      ExtraBallLight.State = 2
      DMDFlush
      DMD RightLine(0, ""), CenterLine(1, ""), 4, eNone, eBlink, eBlinkFast, 1000, True, "WellDone"
    Else 
      DMD CenterLine(0, "_" & Credits), CenterLine(1, ExtraBallCount& " COLLECT EXTRABALL "), 0, eNone, eBlink, eNone, 800, True, ""
      DMD_DisplaySceneEx "DMD1.png", "COLLECT FOR", 15, 1, "EXTRABALL: " & ExtraBallCount , -1, -1, UltraDMD_Animation_None, 800, UltraDMD_Animation_None
   End If

    If ExtraBallLight.State = 2 Then
        AwardExtraBallBlinkLightOn
        BackFlashEffect 3
        BackFlashEffect 4
        LightEffect 1
    End If

    ' remember last trigger hit by the ball
    LastSwitchHit = "Sw26"
End Sub


Sub Sw26a_hit
    If  ComboblinkLightTimer.enabled = True Then 'combo
        DMDFlush
        DMD RightLine(0, ""), CenterLine(1, ""), 40, eNone, eBlink, eBlinkFast, 500, True, "Toasty"
        DMD "_", CenterLine(1, "COMBO! " & FormatScore("20000") ), 0, eNone, eBlinkFast, eNone, 500, True, ""
        DMD_DisplaySceneEx "DMD1.png", "COMBO! 20000", 14, 2, "" , -1, -1, UltraDMD_Animation_None, 800, UltraDMD_Animation_None
        AddScore 20000
        LightEffect 1
        BackFlashEffect 1
        LightSeqCombo.Play SeqBlinking, , 5, 100
    End If
End Sub

'Right Ramp
Sub sw24_Hit
    PlaySound "fx_rampL", 0, 1, pan(ActiveBall)
    If Tilted Then Exit Sub
    AddScore 2000
    SetFlash 1, 2
	DOF 206, DOFPulse
'    startB2S(18)
    vpmTimer.AddTimer 100, "ComboblinkLight"
 '   LightDragon.BlinkInterval = 160
    ExtraBallCount = ExtraBallCount + 1

If bMultiBallMode Then
 If bSuperJackpotMode = false And LJ4.state = 2 Then
     LJ4.state = 1
     AwardJackpot
     ChectSuperJackpot()
   Else
 If LJ4.state = 1 Then
     ChectSuperJackpot()
  End If
 End If
End If

   If ExtraBallCount = 8 or ExtraBallCount = 20 Or ExtraBallCount = 40 Then
      AwardExtraBall()
      ExtraBallLight.State = 2
      DMDFlush
      DMD RightLine(0, ""), CenterLine(1, ""), 4, eNone, eBlink, eBlinkFast, 1000, True, "WellDone" 
    Else 
      DMD CenterLine(0, "_" & Credits), CenterLine(1, ExtraBallCount& " COLLECT EXTRABALL "), 0, eNone, eBlink, eNone, 800, True, ""
      DMD_DisplaySceneEx "DMD1.png", "COLLECT FOR", 15, 1, "EXTRABALL: " & ExtraBallCount , -1, -1, UltraDMD_Animation_None, 800, UltraDMD_Animation_None
   End If

    If ExtraBallLight.State = 2 Then
        AwardExtraBallBlinkLightOn
        BackFlashEffect 3
        BackFlashEffect 4
        LightEffect 1
    End If


    ' remember last trigger hit by the ball
    LastSwitchHit = "Sw24"
End Sub

Sub Sw24z_hit
    If  ComboblinkLightTimer.enabled = True Then 'combo
        DMDFlush
        DMD RightLine(0, ""), CenterLine(1, ""), 40, eNone, eBlink, eBlinkFast, 500, True, "Toasty"
        DMD "_", CenterLine(1, "COMBO! " & FormatScore("20000") ), 0, eNone, eBlinkFast, eNone, 500, True, ""
        DMD_DisplaySceneEx "DMD1.png", "COMBO! 20000", 14, 2, "" , -1, -1, UltraDMD_Animation_None, 800, UltraDMD_Animation_None
        AddScore 20000
        LightEffect 1
        BackFlashEffect 1
        LightSeqCombo.Play SeqBlinking, , 5, 100
    End If
end Sub

Sub Sw27_hit
    PlaySound SoundFXDOF("fx_rubber",116,DOFPulse,DOFTargets), 0, 1, pan(ActiveBall)
    If Tilted Then Exit Sub
    PlaySoundEffect
    Addscore 20000
  If bMultiBallMode Then Exit Sub
    LockLight1.State = 1
    CheckLock
End Sub  

Sub Sw28_hit
    PlaySound SoundFXDOF("fx_rubber",114,DOFPulse,DOFTargets), 0, 1, pan(ActiveBall)
    If Tilted Then Exit Sub
    PlaySoundEffect
    Addscore 20000
  If bMultiBallMode Then Exit Sub
    LockLight2.State = 1
    CheckLock
End Sub

Sub Sw29_hit
    PlaySound SoundFXDOF("fx_rubber",115,DOFPulse,DOFTargets), 0, 1, pan(ActiveBall)
    If Tilted Then Exit Sub
    PlaySoundEffect
    Addscore 20000
  If bMultiBallMode Then Exit Sub
    LockLight3.State = 1
    CheckLock
End Sub

Sub Sw30_Spin()
    PlaySound "Kahn_hit"', 0, 1, pan(ActiveBall)
    If Tilted Then Exit Sub
    If  LastSwitchHit = "Sw19" or LastSwitchHit = "Sw18" and ComboblinkLightTimer.enabled = True Then 'combo
        DMDFlush
        DMD RightLine(0, ""), CenterLine(1, ""), 40, eNone, eBlink, eBlinkFast, 500, True, "Toasty"
        DMD "_", CenterLine(1, "COMBO! " & FormatScore("20000") ), 0, eNone, eBlinkFast, eNone, 500, True, ""
        DMD_DisplaySceneEx "DMD1.png", "COMBO! 20000", 14, 2, "" , -1, -1, UltraDMD_Animation_None, 800, UltraDMD_Animation_None
        AddScore 20000
        LightEffect 1
        BackFlashEffect 1
        LightSeqCombo.Play SeqBlinking, , 5, 100
    End If
        LastSwitchHit = "Sw30"
End Sub
 
Sub CheckLock()
    If bLockEnabled Then Exit Sub
    If LockLight1.State + LockLight2.State + LockLight3.State = 3 Then
        PlaySound "fx_diverter"
        Post.Isdropped = 1
        LockLight.State = 2
        bLockEnabled = True
        MultiLight1.state = 2
        MultiLight2.state = 2
        MultiLight3.state = 2
        DMD "", "", 31, eNone, eNone, eBlinkFast, 1000, True, ""
        DMD_DisplaySceneEx "DMD1.png", "  LOCK IS LIT  ", 14, 2, "" , -1, -1, UltraDMD_Animation_None, 1000, UltraDMD_Animation_None
    End If
End Sub


Sub CheckResetMultiball()

  If bLockEnabled = False Then
     LockLight.State = 0
     MultiLight1.state = 0
     MultiLight2.state = 0
     MultiLight3.state = 0
     LockLight1.State = 2
     LockLight2.State = 2
     LockLight3.State = 2
    Else
     PlaySound "fx_diverter"
     Post.Isdropped = 1
     LockLight.State = 2
     MultiLight1.state = 2
     MultiLight2.state = 2
     MultiLight3.state = 2
     LockLight1.State = 1
     LockLight2.State = 1
     LockLight3.State = 1

  If LockedBalls = 1 Then
     MultiLight1.state = 1
     Else
     MultiLight1.state = 2
  End If

  If LockedBalls = 2 Then
     MultiLight2.state = 1
     MultiLight1.state = 1
     Else
     MultiLight2.state = 2
  End If

  If LockedBalls = 3 Then
     MultiLight3.state = 1
     MultiLight2.state = 1
     MultiLight1.state = 1
     Else
     MultiLight3.state = 2
  End If
  End If


  If MKIC = 1 Then
     SetFlash 3, 1
   Else
     SetFlash 3, 0
  End If           

  If MKIIIC = 1 Then
     SetFlash 5, 1
   Else
     SetFlash 5, 0
  End If

  If MKIIC = 1 Then
     SetFlash 4, 1
   Else
     SetFlash 4, 0
  End If

'    ShaoMultiball = 0
    KinMultiball = 0
    UMultiball = 0
    CharacterMode = False
End Sub

Sub ChectSuperJackpot()

    If bSuperJackpotMode = True And CharacterMode = False Then
        AwardSuperJackpot
    End If

    If bSuperJackpotMode = True and CharacterMode = True Then
        CharactersDefeated()
    End If

    If bSuperJackpotMode = True And CharacterMode = False Then
        AwardSuperJackpot
    End If

    If bSuperJackpotMode = True and CharacterMode = True Then
        CharactersDefeated()
    End If

      'Turn On the Super Jackpot lights
    If (LJ1.state = 1) and (LJ2.state = 1) And (LJ3.state = 1) And (LJ4.state = 1) Then
        CenterRampLight.state = 2
        RightRampLight.state = 2
        PlaySound "Finish_Him"
        LightSeqSuperJackpot.UpdateInterval = 25
        LightSeqSuperJackpot.Play SeqDownOn, 2000, 1
        bSuperJackpotMode = True  
       If CenterRampLight.state = 2 Or RightRampLight.state = 2 then
          FlashForms FlasherLeftRamp, 1000, 40, 0:DOF 201, DOFPulse
          FlashForms FlasherRightRamp, 1000, 40, 0:DOF 204, DOFPulse
          LJ1.state = 2
          LJ4.state = 2
       End If 
    End If
End Sub

sub ComboblinkLight(dummy)
    l16.BlinkInterval = 150
    l21.BlinkInterval = 150
    l20.BlinkInterval = 150
    l16.state=2
    l20.state=2
    l21.state=2
    ComboblinkLightTimer.Enabled = True
end sub

sub ComboblinkLightTimer_timer()
    l16.state=0
    l20.state=0
    l21.state=0
    LastSwitchHit = "Sw28"
    me.enabled=0
end sub

'****************
'* Animation Toys
'****************

'Spin Kung Lao
Dim cBall
KungInit

Sub KungInit
    Set cBall = ckicker.createball
    ckicker.Kick 0, 0
End Sub


Sub KungLaoTrompo()
    KungLao
End Sub

Sub KungLao
    cball.vely = 150 
End Sub



'**********
' Hologram
'**********


Dim Holostep:Holostep = 1
Sub HologramAnim_Timer()
    HologramRampF.visible = 1
    Holostep = Holostep * 1
Select Case Holostep
Case 1:HologramRampF.imageA = "Ramdon1" :Holostep = 2
Case 2:HologramRampF.imageA = "Ramdon2":Holostep = 3
Case 3:HologramRampF.imageA = "Ramdon3":Holostep = 4
Case 4:HologramRampF.imageA = "Ramdon4":Holostep = 5
Case 5:HologramRampF.imageA = "Ramdon5":Holostep = 6
Case 6:HologramRampF.imageA = "Ramdon6":Holostep = 7
Case 7:HologramRampF.imageA = "Ramdon7":Holostep = 8
Case 8:HologramRampF.imageA = "Ramdon8":Holostep = 9
Case 9:HologramRampF.imageA = "Ramdon9":Holostep = 10
Case 10:HologramRampF.imageA = "Ramdon10":Holostep = 11
Case 11:HologramRampF.imageA = "Ramdon11":Holostep = 12
Case 12:HologramRampF.imageA = "Ramdon12":Holostep = 13
Case 13:HologramRampF.imageA = "Ramdon13":Holostep = 14
Case 14:HologramRampF.imageA = "Ramdon14":Holostep = 15
Case 15:HologramRampF.imageA = "Ramdon15":Holostep = 16
Case 16:HologramRampF.imageA = "Ramdon16":Holostep = 17
Case 17:HologramRampF.imageA = "Ramdon17":Holostep = 18
Case 18:HologramRampF.imageA = "Ramdon18":Holostep = 19
Case 19:HologramRampF.imageA = "Ramdon19":Holostep = 20
Case 20:HologramRampF.imageA = "Ramdon20":Holostep = 1
End Select
 
end Sub
Sub HologramRampFOFF()
    HologramRampF.visible = 0
end Sub


Sub IntroAnim()
HologramAnim2.enabled = True
End Sub

Dim Holostep2
Sub HologramAnim2_Timer()
    HologramAnim2.interval = 160
    HologramRampF.visible = 1
    Holostep2 = Holostep2 * 1
Select Case Holostep2
Case 1:HologramRampF.imageA = "Intro000" :Holostep2 = 2
Case 2:HologramRampF.imageA = "Intro001":Holostep2 = 3
Case 3:HologramRampF.imageA = "Intro002":Holostep2 = 4
Case 4:HologramRampF.imageA = "Intro003":Holostep2 = 5
Case 5:HologramRampF.imageA = "Intro004":Holostep2 = 6
Case 6:HologramRampF.imageA = "Intro005":Holostep2 = 7
Case 7:HologramRampF.imageA = "Intro006":Holostep2 = 8
Case 8:HologramRampF.imageA = "Intro007":Holostep2 = 9
Case 9:HologramRampF.imageA = "Intro008":Holostep2 = 10
Case 10:HologramRampF.imageA = "Intro009":Holostep2 = 11
Case 11:HologramRampF.imageA = "Intro010":Holostep2 = 12
Case 12:HologramRampF.imageA = "Intro011":Holostep2 = 13
Case 13:HologramRampF.imageA = "Intro012":Holostep2 = 14
Case 14:HologramRampF.imageA = "Intro013":Holostep2 = 15
Case 15:HologramRampF.imageA = "Intro014":Holostep2 = 16
Case 16:HologramRampF.imageA = "Intro015":Holostep2 = 17
Case 17:HologramRampF.imageA = "Intro016":Holostep2 = 18
Case 18:HologramRampF.imageA = "Intro017":Holostep2 = 19
Case 19:HologramRampF.imageA = "Intro018":Holostep2 = 20
Case 20:HologramRampF.imageA = "Intro019":Holostep2 = 21

Case 21:HologramRampF.imageA = "Intro020" :Holostep2 = 22
Case 22:HologramRampF.imageA = "Intro021":Holostep2 = 23
Case 23:HologramRampF.imageA = "Intro022":Holostep2 = 24
Case 24:HologramRampF.imageA = "Intro023":Holostep2 = 25
Case 25:HologramRampF.imageA = "Intro024":Holostep2 = 26
Case 26:HologramRampF.imageA = "Intro025":Holostep2 = 27
Case 27:HologramRampF.imageA = "Intro026":Holostep2 = 28
Case 28:HologramRampF.imageA = "Intro027":Holostep2 = 29
Case 29:HologramRampF.imageA = "Intro028":Holostep2 = 30
Case 30:HologramRampF.imageA = "Intro029":Holostep2 = 31
Case 31:HologramRampF.imageA = "Intro030":Holostep2 = 32
Case 32:HologramRampF.imageA = "Intro031":Holostep2 = 33
Case 33:HologramRampF.imageA = "Intro032":Holostep2 = 34
Case 34:HologramRampF.imageA = "Intro033":Holostep2 = 35
Case 35:HologramRampF.imageA = "Intro034":Holostep2 = 36
Case 36:HologramRampF.imageA = "Intro035":Holostep2 = 37
Case 37:HologramRampF.imageA = "Intro036":Holostep2 = 38
Case 38:HologramRampF.imageA = "Intro037":Holostep2 = 39
Case 39:HologramRampF.imageA = "Intro038":Holostep2 = 40
Case 40:HologramRampF.imageA = "Intro039":Holostep2 = 41

Case 41:HologramRampF.imageA = "Intro040" :Holostep2 = 42
Case 42:HologramRampF.imageA = "Intro041":Holostep2 = 43
Case 43:HologramRampF.imageA = "Intro042":Holostep2 = 44
Case 44:HologramRampF.imageA = "Intro043":Holostep2 = 45
Case 45:HologramRampF.imageA = "Intro044":Holostep2 = 46
Case 46:HologramRampF.imageA = "Intro045":Holostep2 = 47
Case 47:HologramRampF.imageA = "Intro046":Holostep2 = 48
Case 48:HologramRampF.imageA = "Intro047":Holostep2 = 49
Case 49:HologramRampF.imageA = "Intro048":Holostep2 = 50
Case 50:HologramRampF.imageA = "Intro049":Holostep2 = 51
Case 51:HologramRampF.imageA = "Intro050":Holostep2 = 52
Case 52:HologramRampF.imageA = "Intro051":Holostep2 = 53
Case 53:HologramRampF.imageA = "Intro052":Holostep2 = 54
Case 54:HologramRampF.imageA = "Intro053":Holostep2 = 55
Case 55:HologramRampF.imageA = "Intro054":Holostep2 = 56
Case 56:HologramRampF.imageA = "Intro055":Holostep2 = 57
Case 57:HologramRampF.imageA = "Intro056":Holostep2 = 58
Case 58:HologramRampF.imageA = "Intro057":Holostep2 = 59
Case 59:HologramRampF.imageA = "Intro058":Holostep2 = 60
Case 60:HologramRampF.imageA = "Intro059":Holostep2 = 61

Case 61:HologramRampF.imageA = "Intro060" :Holostep2 = 62
Case 62:HologramRampF.imageA = "Intro061":Holostep2 = 63
Case 63:HologramRampF.imageA = "Intro062":Holostep2 = 64
Case 64:HologramRampF.imageA = "Intro063":Holostep2 = 65
Case 65:HologramRampF.imageA = "Intro064":Holostep2 = 66
Case 66:HologramRampF.imageA = "Intro065":Holostep2 = 67
Case 67:HologramRampF.imageA = "Intro066":Holostep2 = 68
Case 68:HologramRampF.imageA = "Intro067":Holostep2 = 69
Case 69:HologramRampF.imageA = "Intro068":Holostep2 = 70
Case 70:HologramRampF.imageA = "Intro069":Holostep2 = 71
Case 71:HologramRampF.imageA = "Intro070":Holostep2 = 72
Case 72:HologramRampF.imageA = "Intro071":Holostep2 = 73
Case 73:HologramRampF.imageA = "Intro072":Holostep2 = 74
Case 74:HologramRampF.imageA = "Intro073":Holostep2 = 75
Case 75:HologramRampF.imageA = "Intro074":Holostep2 = 76
Case 76:HologramRampF.imageA = "Intro075":Holostep2 = 77
Case 77:HologramRampF.imageA = "Intro076":Holostep2 = 78
Case 78:HologramRampF.imageA = "Intro077":Holostep2 = 79
Case 79:HologramRampF.imageA = "Intro078":Holostep2 = 80
Case 80:HologramRampF.imageA = "Intro079":Holostep2 = 81

Case 81:HologramRampF.imageA = "Intro080" :Holostep2 = 82
Case 82:HologramRampF.imageA = "Intro081":Holostep2 = 83
Case 83:HologramRampF.imageA = "Intro082":Holostep2 = 84
Case 84:HologramRampF.imageA = "Intro083":Holostep2 = 85
Case 85:HologramRampF.imageA = "Intro084":Holostep2 = 86
Case 86:HologramRampF.imageA = "Intro085":Holostep2 = 87
Case 87:HologramRampF.imageA = "Intro086":Holostep2 = 88
Case 88:HologramRampF.imageA = "Intro087":Holostep2 = 89
Case 89:HologramRampF.imageA = "Intro088":Holostep2 = 90
Case 90:HologramRampF.imageA = "Intro089":Holostep2 = 91
Case 91:HologramRampF.imageA = "Intro090":Holostep2 = 92
Case 92:HologramRampF.imageA = "Intro091":Holostep2 = 93
Case 93:HologramRampF.imageA = "Intro092":Holostep2 = 94
Case 94:HologramRampF.imageA = "Intro093":Holostep2 = 95
Case 95:HologramRampF.imageA = "Intro094":Holostep2 = 96
Case 96:HologramRampF.imageA = "Intro095":Holostep2 = 97
Case 97:HologramRampF.imageA = "Intro096":Holostep2 = 98
Case 98:HologramRampF.imageA = "Intro097":Holostep2 = 99
Case 99:HologramRampF.imageA = "Intro098":Holostep2 = 100
Case 100:HologramRampF.imageA = "Intro099":Holostep2 = 101

Case 101:HologramRampF.imageA = "Intro100" :Holostep2 = 102
Case 102:HologramRampF.imageA = "Intro101":Holostep2 = 103
Case 103:HologramRampF.imageA = "Intro102":Holostep2 = 104
Case 104:HologramRampF.imageA = "Intro103":Holostep2 = 105
Case 105:HologramRampF.imageA = "Intro104":Holostep2 = 106
Case 106:HologramRampF.imageA = "Intro105":Holostep2 = 107
Case 107:HologramRampF.imageA = "Intro106":Holostep2 = 108
Case 108:HologramRampF.imageA = "Intro107":Holostep2 = 109
Case 109:HologramRampF.imageA = "Intro108":Holostep2 = 110
Case 110:HologramRampF.imageA = "Intro109":Holostep2 = 111
Case 111:HologramRampF.imageA = "Intro110":Holostep2 = 112
Case 112:HologramRampF.imageA = "Intro111":Holostep2 = 113
Case 113:HologramRampF.imageA = "Intro112":Holostep2 = 114
Case 114:HologramRampF.imageA = "Intro113":Holostep2 = 115
Case 115:HologramRampF.imageA = "Intro114":Holostep2 = 116
Case 116:HologramRampF.imageA = "Intro115":Holostep2 = 117
Case 117:HologramRampF.imageA = "Intro116":Holostep2 = 118
Case 118:HologramRampF.imageA = "Intro117":Holostep2 = 119
Case 119:HologramRampF.imageA = "Intro118":Holostep2 = 120
Case 120:HologramRampF.imageA = "Intro119":Holostep2 = 121
Case 121:HologramRampF.imageA = "Intro120":Holostep2 = 122


Case 122:HologramRampF.imageA = "Intro121":Holostep2 = 123
Case 123:HologramRampF.imageA = "Intro122":Holostep2 = 124
Case 124:HologramRampF.imageA = "Intro123":Holostep2 = 125
Case 125:HologramRampF.imageA = "Intro124":Holostep2 = 126
Case 126:HologramRampF.imageA = "Intro125":Holostep2 = 127
Case 127:HologramRampF.imageA = "Intro126":Holostep2 = 128
Case 128:HologramRampF.imageA = "Intro127":Holostep2 = 129
Case 129:HologramRampF.imageA = "Intro128":Holostep2 = 130
Case 130:HologramRampF.imageA = "Intro129":Holostep2 = 131
Case 131:HologramRampF.imageA = "Intro130":Holostep2 = 132
Case 132:HologramRampF.imageA = "Intro131":Holostep2 = 133
Case 133:HologramRampF.imageA = "Intro132":Holostep2 = 134
Case 134:HologramRampF.imageA = "Intro133":Holostep2 = 135
Case 135:HologramRampF.imageA = "Intro134":Holostep2 = 136
Case 136:HologramRampF.imageA = "Intro135":Holostep2 = 137
Case 137:HologramRampF.imageA = "Intro136":Holostep2 = 138
Case 138:HologramRampF.imageA = "Intro137":Holostep2 = 139
Case 139:HologramRampF.imageA = "Intro138":Holostep2 = 140
Case 140:HologramRampF.imageA = "Intro139":Holostep2 = 141
Case 141:HologramRampF.imageA = "Intro140":Holostep2 = 142


Case 142:HologramRampF.imageA = "Intro141":Holostep2 = 143
Case 143:HologramRampF.imageA = "Intro142":Holostep2 = 144
Case 144:HologramRampF.imageA = "Intro143":Holostep2 = 145
Case 145:HologramRampF.imageA = "Intro144":Holostep2 = 146
Case 146:HologramRampF.imageA = "Intro145":Holostep2 = 147
Case 147:HologramRampF.imageA = "Intro146":Holostep2 = 148
Case 148:HologramRampF.imageA = "Intro147":Holostep2 = 149
Case 149:HologramRampF.imageA = "Intro148":Holostep2 = 150
Case 150:HologramRampF.imageA = "Intro149":Holostep2 = 151
Case 151:HologramRampF.imageA = "Intro150":Holostep2 = 152
Case 152:HologramRampF.imageA = "Intro151":Holostep2 = 153
Case 153:HologramRampF.imageA = "Intro152":Holostep2 = 154
Case 154:HologramRampF.imageA = "Intro153":Holostep2 = 155
Case 155:HologramRampF.imageA = "Intro154":Holostep2 = 156
Case 156:HologramRampF.imageA = "Intro155":Holostep2 = 157
Case 157:HologramRampF.imageA = "Intro156":Holostep2 = 158
Case 158:HologramRampF.imageA = "Intro157":Holostep2 = 159
Case 159:HologramRampF.imageA = "Intro158":Holostep2 = 160
Case 160:HologramRampF.imageA = "Intro159":Holostep2 = 161
Case 161:HologramRampF.imageA = "Intro160":Holostep2 = 162

Case 162:HologramRampF.imageA = "Intro161":Holostep2 = 163
Case 163:HologramRampF.imageA = "Intro162":Holostep2 = 164
Case 164:HologramRampF.imageA = "Intro163":Holostep2 = 165
Case 165:HologramRampF.imageA = "Intro164":Holostep2 = 166
Case 166:HologramRampF.imageA = "Intro165":Holostep2 = 167
Case 167:HologramRampF.imageA = "Intro166":Holostep2 = 168
Case 168:HologramRampF.imageA = "Intro167":Holostep2 = 169
Case 169:HologramRampF.imageA = "Intro168":Holostep2 = 170
Case 170:HologramRampF.imageA = "Intro169":Holostep2 = 171
Case 171:HologramRampF.imageA = "Intro170":Holostep2 = 172
Case 172:HologramRampF.imageA = "Intro171":Holostep2 = 173
Case 173:HologramRampF.imageA = "Intro172":Holostep2 = 174
Case 174:HologramRampF.imageA = "Intro173":Holostep2 = 0
         HologramRampF.visible = 0: Me.enabled = False
End Select

End Sub


Sub IntroAnimOff()
HologramAnim2.enabled = False
HologramRampF.visible = 0
End Sub



'***************
' Random Select
'***************

Dim SlotPos
Sub StartSlotmachine() ' uses the HolePos variable
    Dim i
    HologramAnim.enabled = 0
    HologramAnim.enabled = 1
	DOF 135, DOFOn
    SlotPos = 8
    DMDFlush
    For i = SlotPos to SlotPos * 2 
        DMD "", "", i, eNone, eNone, eNone, 2, False, "fx_spinner"
        CharactersRamdonUltraDMD.enabled = 1
    Next 
    vpmtimer.AddTimer 1500, "GiveSlotAward '"
    GiEffect 1
End Sub

Sub GiveSlotAward()
    Dim tmp
    HologramAnim.enabled = 1
    HologramAnim.enabled = 0
	DOF 135, DOFOff
    RamdonUltraDMDOff
    DMDFlush
    tmp = INT(SlotPos + RND * 16)
    debug.print "slot pos." & tmp
    DMD "", "", tmp, eNone, eNone, eBlinkFast, 800, True, ""
  
    vpmtimer.AddTimer 2000, "HologramRampFOFF '"
    Select Case tmp
            
        Case 18:     If ScorpionD = 0 Then
                     DMDFlush
                     debug.print "slot pos." & tmp
                     DMD RightLine(0, ""), CenterLine(1, ""), 18, eNone, eBlink, eNone, 1000, True, "Scorpion"
                     DMD "_", CenterLine(1, "AWARD  " & FormatScore("2500") ), 33, eNone, eBlinkFast, eNone, 500, True, ""
                     DMD_DisplaySceneEx "DMD1.png", "Scorpion" , 15, 4, "AWARD: 2500" , -1, -1, UltraDMD_Animation_None, 1200, UltraDMD_Animation_None
                     HologramRampF.imageA = "RamdonScorpion"
                     AddScore 2500 
                     Light11.state = 0: If B2SOn Then Controller.B2SSetData 11, 0
                     ScorpionD = 1
                     CraractersDefeated()
                     Else
                     StartSlotmachine
                     End If

        Case 19:     if RaidenD = 0 then
                     DMDFlush
                     debug.print "slot pos." & tmp
                     DMD RightLine(0, ""), CenterLine(1, ""), 19, eNone, eBlink, eNone, 1000, True, "RAIDEN"
                     DMD "_", CenterLine(1, "AWARD  " & FormatScore("2500") ), 33, eNone, eBlinkFast, eNone, 500, True, ""
                     DMD_DisplaySceneEx "DMD1.png", "RAIDEN" , 15, 4, "AWARD: 2500" , -1, -1, UltraDMD_Animation_None, 1200, UltraDMD_Animation_None
                     HologramRampF.imageA = "RamdonRaiden"
                     AddScore 2500
                     Light12.state = 0: If B2SOn Then Controller.B2SSetData 12, 0
                     RaidenD = 1
                     CraractersDefeated()
                     Else
                     StartSlotmachine
                     End If
                     
        Case 8:     If LiuKangD = 0 Then
                     DMDFlush
                     debug.print "slot pos." & tmp
                     DMD RightLine(0, ""), CenterLine(1, ""), 8, eNone, eBlink, eNone, 1000, True, "Liu_Kang"
                     DMD "_", CenterLine(1, "AWARD  " & FormatScore("2500") ), 33, eNone, eBlinkFast, eNone, 500, True, ""
                     DMD_DisplaySceneEx "DMD1.png", "Liu Kang" , 15, 4, "AWARD: 2500"  , -1, -1, UltraDMD_Animation_None, 1200, UltraDMD_Animation_None
                     HologramRampF.imageA = "RamdonLiu"
                     AddScore 2500 
                     Light1.state = 0: If B2SOn Then Controller.B2SSetData 1, 0
                     LiuKangD = 1
                     CraractersDefeated()
                     Else
                     StartSlotmachine
                     End If

        Case 9:      If KungLaoD = 0 then
                     DMDFlush
                     debug.print "slot pos." & tmp
                     DMD RightLine(0, ""), CenterLine(1, ""), 9, eNone, eBlink, eNone, 1000, True, "KUNG_LAO"
                     DMD "_", CenterLine(1, "AWARD  " & FormatScore("2500") ), 33, eNone, eBlinkFast, eNone, 500, True, ""
                     DMD_DisplaySceneEx "DMD1.png", "KUNG LAO" , 15, 4, "AWARD: 2500"  , -1, -1, UltraDMD_Animation_None, 1200, UltraDMD_Animation_None
                     HologramRampF.imageA = "RamdonKung"
                     AddScore 2500
                     Light2.state = 0: If B2SOn Then Controller.B2SSetData 2, 0 
                     KungLaoD = 1
                     CraractersDefeated()
                     Else
                     StartSlotmachine
                     End If

        Case 15:     If JaxD = 0 Then
                     DMDFlush
                     debug.print "slot pos." & tmp
                     DMD RightLine(0, ""), CenterLine(1, ""), 15, eNone, eBlink, eNone, 1000, True, "JAX"
                     DMD "_", CenterLine(1, "AWARD  " & FormatScore("2500") ), 33, eNone, eBlinkFast, eNone, 500, True, ""
                     DMD_DisplaySceneEx "DMD1.png", "Jax" , 15, 4, "AWARD: 2500"  , -1, -1, UltraDMD_Animation_None, 1200, UltraDMD_Animation_None
                     HologramRampF.imageA = "RamdonJax"
                     Light8.state = 0: If B2SOn Then Controller.B2SSetData 8, 0
                     AddScore 2500
                     JaxD = 1
                     CraractersDefeated()
                     Else
                     StartSlotmachine
                     End If

        Case 16:     If MillenaD = 0 Then
                     DMDFlush
                     debug.print "slot pos." & tmp
                     DMD RightLine(0, ""), CenterLine(1, ""), 16, eNone, eBlink, eNone, 1000, True, "MILEENA"
                     DMD "_", CenterLine(1, "AWARD  " & FormatScore("2500") ), 33, eNone, eBlinkFast, eNone, 500, True, ""
                     DMD_DisplaySceneEx "DMD1.png", "Millena" , 15, 4, "AWARD: 2500"  , -1, -1, UltraDMD_Animation_None, 1200, UltraDMD_Animation_None
                     HologramRampF.imageA = "RamdonMillena"
                     AddScore 2500
                     Light9.state = 0: If B2SOn Then Controller.B2SSetData 9, 0
                     MillenaD = 1
                     CraractersDefeated()
                     Else
                     StartSlotmachine
                     End If

        Case 14:     If KitanaD = 0 Then
                     DMDFlush
                     debug.print "slot pos." & tmp
                     DMD RightLine(0, ""), CenterLine(1, ""), 14, eNone, eBlink, eNone, 1000, True, "KITANA"
                     DMD "_", CenterLine(1, "AWARD  " & FormatScore("2500") ), 33, eNone, eBlinkFast, eNone, 500, True, ""
                     DMD_DisplaySceneEx "DMD1.png", "Kitana" , 15, 4, "AWARD: 2500"  , -1, -1, UltraDMD_Animation_None, 1200, UltraDMD_Animation_None
                     HologramRampF.imageA = "RamdonKitana"
                     AddScore 2500
                     Light7.state = 0: If B2SOn Then Controller.B2SSetData 7, 0
                     KitanaD = 1
                     CraractersDefeated()
                     Else
                     StartSlotmachine
                     End If

        Case 20:     If SmokeD = 0 Then
                     DMDFlush
                     debug.print "slot pos." & tmp
                     DMD RightLine(0, ""), CenterLine(1, ""), 20, eNone, eBlink, eNone, 1000, True, "SMOKE"
                     DMD "_", CenterLine(1, "AWARD  " & FormatScore("4000") ), 33, eNone, eBlinkFast, eNone, 500, True, ""
                     DMD_DisplaySceneEx "DMD1.png", "Smoke" , 15, 4, "AWARD: 2500"  , -1, -1, UltraDMD_Animation_None, 1200, UltraDMD_Animation_None
                     HologramRampF.imageA = "RamdonSmoke"
                     AddScore 4000
                     Light13.state = 0: If B2SOn Then Controller.B2SSetData 13, 0
                     SmokeD = 1
                     UMultiball = 1
                     CraractersDefeated()
                     Else
                     StartSlotmachine
                     End If

        Case 22:     If KintaroD = 1 Then
                     DMDFlush
                     debug.print "slot pos." & tmp
                     DMD RightLine(0, ""), CenterLine(1, ""), 22, eNone, eBlink, eNone, 1000, True, "KINTARO"
                     DMD "_", CenterLine(1, "AWARD  " & FormatScore("5000") ), 33, eNone, eBlinkFast, eNone, 500, True, ""
                     DMD_DisplaySceneEx "DMD1.png", "Kintaro" , 15, 4, "AWARD: 2500"  , -1, -1, UltraDMD_Animation_None, 1200, UltraDMD_Animation_None
                     HologramRampF.imageA = "RamdonKintaro"
                     AddScore 5000
                     Light15.state = 0: If B2SOn Then Controller.B2SSetData 15, 0
                     KintaroD = 2
                     KinMultiball = 1
                     CraractersDefeated()
                     Else
                     StartSlotmachine
                     End If

        Case 23:     If ShaoKahnD = 1 Then
                     DMDFlush
                     debug.print "slot pos." & tmp
                     DMD RightLine(0, ""), CenterLine(1, ""), 23, eNone, eBlink, eNone, 1000, True, "I_Am_Shao_Kahn"
                     DMD "_", CenterLine(1, "AWARD  " & FormatScore("7000") ), 33, eNone, eBlinkFast, eNone, 500, True, ""
                     DMD_DisplaySceneEx "DMD1.png", "Shao Kahn" , 15, 4, "AWARD: 2500"  , -1, -1, UltraDMD_Animation_None, 1200, UltraDMD_Animation_None
                     HologramRampF.imageA = "RamdonShao"
                     AddScore 7000
                     Light16.state = 0: If B2SOn Then Controller.B2SSetData 16, 0
                     ShaoKahnD = 2
                     vpmtimer.AddTimer 1000, "CraractersDefeated '"
               '      CraractersDefeated()
                     Else
                     StartSlotmachine
                     End If

        Case 12:     If SubZeroD = 0 Then
                     DMDFlush
                     debug.print "slot pos." & tmp
                     DMD RightLine(0, ""), CenterLine(1, ""), 12, eNone, eBlink, eNone, 1000, True, "SUB_ZERO"
                     DMD "_", CenterLine(1, "AWARD  " & FormatScore("2500") ), 33, eNone, eBlinkFast, eNone, 500, True, ""
                     DMD_DisplaySceneEx "DMD1.png", "Sub-Zero" , 15, 4, "AWARD: 2500" , -1, -1, UltraDMD_Animation_None, 1200, UltraDMD_Animation_None
                     HologramRampF.imageA = "RamdonSub"
                     AddScore 2500
                     Light5.state = 0: If B2SOn Then Controller.B2SSetData 5, 0
                     SubZeroD = 1
                     CraractersDefeated()
                     Else
                     StartSlotmachine
                     End If

        Case 10:     If JohnnyCajeD = 0 Then
                     DMDFlush
                     debug.print "slot pos." & tmp
                     DMD RightLine(0, ""), CenterLine(1, ""), 10, eNone, eBlink, eNone, 1000, True, "Johnny_Cage"
                     DMD "_", CenterLine(1, "AWARD  " & FormatScore("2500") ), 33, eNone, eBlinkFast, eNone, 500, True, ""
                     DMD_DisplaySceneEx "DMD1.png", "Johnny Caje" , 15, 4, "AWARD: 2500"  , -1, -1, UltraDMD_Animation_None, 1200, UltraDMD_Animation_None
                     HologramRampF.imageA = "RamdonJohnny"
                     AddScore 2500
                     Light3.state = 0: If B2SOn Then Controller.B2SSetData 3, 0
                     JohnnyCajeD = 1
                     CraractersDefeated()
                     Else
                     StartSlotmachine
                     End If

        Case 11:     If ReptileD = 0 Then
                     DMDFlush
                     debug.print "slot pos." & tmp
                     DMD RightLine(0, ""), CenterLine(1, ""), 11, eNone, eBlink, eNone, 1000, True, "REPTILE"
                     DMD "_", CenterLine(1, "AWARD  " & FormatScore("2500") ), 33, eNone, eBlinkFast, eNone, 500, True, ""
                     DMD_DisplaySceneEx "DMD1.png", "Reptile" , 15, 4, "AWARD: 2500" , -1, -1, UltraDMD_Animation_None, 1200, UltraDMD_Animation_None
                     HologramRampF.imageA = "RamdonReptile"
                     AddScore 2500
                     Light4.state = 0: If B2SOn Then Controller.B2SSetData 4, 0
                     ReptileD = 1
                     CraractersDefeated()
                     Else
                     StartSlotmachine
                     End If

        Case 17:     If BarakaD = 0 Then
                     DMDFlush
                     debug.print "slot pos." & tmp
                     DMD RightLine(0, ""), CenterLine(1, ""), 17, eNone, eBlink, eNone, 1000, True, "BARAKA"
                     DMD "_", CenterLine(1, "AWARD  " & FormatScore("2500") ), 33, eNone, eBlinkFast, eNone, 500, True, ""
                     DMD_DisplaySceneEx "DMD1.png", "Baraka" , 15, 4, "AWARD: 2500" , -1, -1, UltraDMD_Animation_None, 1200, UltraDMD_Animation_None
                     HologramRampF.imageA = "RamdonBaraka"
                     Light10.state = 0: If B2SOn Then Controller.B2SSetData 9, 0
                     AddScore 2500
                     BarakaD = 1
                     CraractersDefeated()                     
                     Else
                     StartSlotmachine
                     End If

        Case 13:     If ShangTsungD = 0 Then
                     DMDFlush
                     debug.print "slot pos." & tmp
                     DMD RightLine(0, ""), CenterLine(1, ""), 13, eNone, eBlink, eNone, 1000, True, "Shang_Tsung"
                     DMD "_", CenterLine(1, "AWARD  " & FormatScore("3000") ), 33, eNone, eBlinkFast, eNone, 500, True, ""
                     DMD_DisplaySceneEx "DMD1.png", "Shang Tsung" , 15, 4,"AWARD: 2500"  , -1, -1, UltraDMD_Animation_None, 1200, UltraDMD_Animation_None
                     HologramRampF.imageA = "RamdonShang"
                     AddScore 3000
                     Light6.state = 0: If B2SOn Then Controller.B2SSetData 6, 0
                     ShangTsungD = 1
                     CraractersDefeated()
                     Else
                     StartSlotmachine
                     End If

        Case 21:     If JadeD = 0 Then
                     DMDFlush
                     debug.print "slot pos." & tmp
                     DMD RightLine(0, ""), CenterLine(1, ""), 21, eNone, eBlink, eNone, 1000, True, "JADE"
                     DMD "_", CenterLine(1, "AWARD  " & FormatScore("4000") ), 33, eNone, eBlinkFast, eNone, 500, True, ""
                     DMD_DisplaySceneEx "DMD1.png", "Jade" , 15, 4, "AWARD: 2500"  , -1, -1, UltraDMD_Animation_None, 1200, UltraDMD_Animation_None
                     HologramRampF.imageA = "RamdonJade"
                     AddScore 4000
                     Light14.state = 0: If B2SOn Then Controller.B2SSetData 14, 0
                     JadeD = 1
                     UMultiball = 2
                     CraractersDefeated()
                     Else
                     StartSlotmachine
                     End If
    End Select

End Sub


Sub CharactersRamdonUltraDMD_Timer
Dim tmp
tmp = INT(SlotPos + RND * 16)
    Select Case tmp

        Case 18:  DMD_DisplaySceneEx "DMD1.png", "Scorpion" , 15, 4, "_" , -1, -1, UltraDMD_Animation_None, 10, UltraDMD_Animation_None

        Case 19:  DMD_DisplaySceneEx "DMD1.png", "RAIDEN" , 15, 4, "_" , -1, -1, UltraDMD_Animation_None, 10, UltraDMD_Animation_None
                     
        Case 8:   DMD_DisplaySceneEx "DMD1.png", "Liu Kang" , 15, 4, "_"  , -1, -1, UltraDMD_Animation_None, 10, UltraDMD_Animation_None

        Case 9:   DMD_DisplaySceneEx "DMD1.png", "KUNG LAO" , 15, 4, "_"  , -1, -1, UltraDMD_Animation_None, 10, UltraDMD_Animation_None

        Case 15:  DMD_DisplaySceneEx "DMD1.png", "Jax" , 15, 4, "_"  , -1, -1, UltraDMD_Animation_None, 10, UltraDMD_Animation_None

        Case 16:  DMD_DisplaySceneEx "DMD1.png", "Millena" , 15, 4, "_"  , -1, -1, UltraDMD_Animation_None, 10, UltraDMD_Animation_None

        Case 14:  DMD_DisplaySceneEx "DMD1.png", "Kitana" , 15, 4, "_"  , -1, -1, UltraDMD_Animation_None, 10, UltraDMD_Animation_None

        Case 20:  DMD_DisplaySceneEx "DMD1.png", "Smoke" , 15, 4, "_"  , -1, -1, UltraDMD_Animation_None, 10, UltraDMD_Animation_None

        Case 22:  DMD_DisplaySceneEx "DMD1.png", "Kintaro" , 15, 4, "_"  , -1, -1, UltraDMD_Animation_None, 10, UltraDMD_Animation_None

        Case 23:  DMD_DisplaySceneEx "DMD1.png", "Shao Kahn" , 15, 4, "_"  , -1, -1, UltraDMD_Animation_None, 10, UltraDMD_Animation_None

        Case 12:  DMD_DisplaySceneEx "DMD1.png", "Sub-Zero" , 15, 4, "_" , -1, -1, UltraDMD_Animation_None, 10, UltraDMD_Animation_None

        Case 10:  DMD_DisplaySceneEx "DMD1.png", "Johnny Caje" , 15, 4, "_"  , -1, -1, UltraDMD_Animation_None, 10, UltraDMD_Animation_None

        Case 11:  DMD_DisplaySceneEx "DMD1.png", "Reptile" , 15, 4, "_" , -1, -1, UltraDMD_Animation_None, 10, UltraDMD_Animation_None

        Case 17:  DMD_DisplaySceneEx "DMD1.png", "Baraka" , 15, 4, "_" , -1, -1, UltraDMD_Animation_None, 10, UltraDMD_Animation_None

        Case 13:  DMD_DisplaySceneEx "DMD1.png", "Shang Tsung" , 15, 4,"_"  , -1, -1, UltraDMD_Animation_None, 10, UltraDMD_Animation_None              

        Case 21:  DMD_DisplaySceneEx "DMD1.png", "Jade" , 15, 4, "_"  , -1, -1, UltraDMD_Animation_None, 10, UltraDMD_Animation_None

    End Select
End Sub

Sub RamdonUltraDMDOff
Dim tmp
    CharactersRamdonUltraDMD.enabled = 0
Dim iScene
    UltraDMD.CancelRendering
    Select Case tmp

        Case 18:   DMD_DisplaySceneEx "DMD1.png", "Scorpion" , 15, 4, "_" , -1, -1, UltraDMD_Animation_None, 10, UltraDMD_Animation_None

        Case 19:   DMD_DisplaySceneEx "DMD1.png", "RAIDEN" , 15, 4, "_" , -1, -1, UltraDMD_Animation_None, 10, UltraDMD_Animation_None
                     
        Case 8:    DMD_DisplaySceneEx "DMD1.png", "Liu Kang" , 15, 4, "_"  , -1, -1, UltraDMD_Animation_None, 10, UltraDMD_Animation_None

        Case 9:    DMD_DisplaySceneEx "DMD1.png", "KUNG LAO" , 15, 4, "_"  , -1, -1, UltraDMD_Animation_None, 10, UltraDMD_Animation_None

        Case 15:   DMD_DisplaySceneEx "DMD1.png", "Jax" , 15, 4, "_"  , -1, -1, UltraDMD_Animation_None, 10, UltraDMD_Animation_None

        Case 16:   DMD_DisplaySceneEx "DMD1.png", "Millena" , 15, 4, "_"  , -1, -1, UltraDMD_Animation_None, 10, UltraDMD_Animation_None

        Case 14:   DMD_DisplaySceneEx "DMD1.png", "Kitana" , 15, 4, "_"  , -1, -1, UltraDMD_Animation_None, 10, UltraDMD_Animation_None

        Case 20:   DMD_DisplaySceneEx "DMD1.png", "Smoke" , 15, 4, "_"  , -1, -1, UltraDMD_Animation_None, 10, UltraDMD_Animation_None

        Case 22:   DMD_DisplaySceneEx "DMD1.png", "Kintaro" , 15, 4, "_"  , -1, -1, UltraDMD_Animation_None, 10, UltraDMD_Animation_None

        Case 23:   DMD_DisplaySceneEx "DMD1.png", "Shao Kahn" , 15, 4, "_"  , -1, -1, UltraDMD_Animation_None, 10, UltraDMD_Animation_None

        Case 12:   DMD_DisplaySceneEx "DMD1.png", "Sub-Zero" , 15, 4, "_" , -1, -1, UltraDMD_Animation_None, 10, UltraDMD_Animation_None

        Case 10:   DMD_DisplaySceneEx "DMD1.png", "Johnny Caje" , 15, 4, "_"  , -1, -1, UltraDMD_Animation_None, 10, UltraDMD_Animation_None

        Case 11:   DMD_DisplaySceneEx "DMD1.png", "Reptile" , 15, 4, "_" , -1, -1, UltraDMD_Animation_None, 10, UltraDMD_Animation_None

        Case 17:   DMD_DisplaySceneEx "DMD1.png", "Baraka" , 15, 4, "_" , -1, -1, UltraDMD_Animation_None, 10, UltraDMD_Animation_None

        Case 13:   DMD_DisplaySceneEx "DMD1.png", "Shang Tsung" , 15, 4,"_"  , -1, -1, UltraDMD_Animation_None, 10, UltraDMD_Animation_None              

        Case 21:   DMD_DisplaySceneEx "DMD1.png", "Jade" , 15, 4, "_"  , -1, -1, UltraDMD_Animation_None, 10, UltraDMD_Animation_None

    End Select

    iScene = (iScene + 1) MOD 16

End Sub
'***********************
' Craracters Defeated
'***********************

Dim DefeatedAll
Dim ScorpionD, RaidenD, LiuKangD, KungLaoD, JaxD, MillenaD, KitanaD, SmokeD, KintaroD, ShaoKahnD, SubZeroD, JohnnyCajeD, ReptileD, BarakaD, ShangTsungD, JadeD


Sub CraractersDefeated()
 If ScorpionD=1 And RaidenD=1 And KungLaoD=1 And JaxD=1 And MillenaD=1 And KitanaD=1 And SmokeD=1 And SubZeroD=1 And JohnnyCajeD=1 And ReptileD=1 And BarakaD=1 And ShangTsungD=1 And JadeD=1 Then
    KintaroD = 1
 End If

if  ShaoKahnD = 2 Then
    StopAllMusic
    DefeatedAll = True 
    DMDFlush
     DMD RightLine(1, ""), RightLine(1, ""), 38, eNone, eNone, eBlink, 3800, True, "End"
     DMD RightLine(1, ""), RightLine(1, ""), 39, eNone, eNone, eBlink, 10000, True, ""
     DMD_DisplaySceneEx "DMD1.png", "ULTIMATE" , 15, 4, "Multiball"  , -1, -1, UltraDMD_Animation_None, 1200, UltraDMD_Animation_None
     AddScore 16000000
     vpmtimer.AddTimer 13800, "CharactersDefeatedMultiball '"
End If

if KinMultiball = 1 Then
   LightSeqMultiball.Play SeqBlinking, , 20, 50
   LightSeqMultiball2.Play SeqBlinking, , 20, 50
   vpmtimer.AddTimer 1500, "KintaroMultiball '"
End If

if UMultiball = 1 Or UMultiball = 2 Then
   LightSeqMultiball.Play SeqBlinking, , 20, 50
   LightSeqMultiball2.Play SeqBlinking, , 20, 50
   vpmtimer.AddTimer 1500, "UnknownMultiball '"
End If

If DefeatedAll = True Then
   vpmtimer.addtimer 13800, "Sw13Hole '"
End If

If DefeatedAll = False Then
   vpmtimer.addtimer 2500, "Sw13Hole '"
End If

End Sub


Sub CharactersDefeatedMultiball()
     ChangeGi "red"
     AddMultiball 3
     EnableBallSaver 15 
     LockLight.State = 0
    'Turn On the Jackpot lights
     LJ1.state = 2
     LJ2.state = 2
     LJ3.state = 2
     LJ4.state = 2
     CenterRampLight.state = 0
     RightRampLight.state = 0
     SetFlash 3, 0
     SetFlash 4, 0
     SetFlash 5, 0
	 If Sw25.Isdropped = 0 Then DOF 114, DOFPulse
     Sw25.isdropped = 1
     l18.State = 0
     Post.Isdropped= 0
     ScorpionD = 0
     RaidenD = 0
     KungLaoD = 0
     JaxD = 0
     MillenaD = 0
     KitanaD = 0
     SmokeD = 0
     KintaroD = 1
     ShaoKahnD = 1
     SubZeroD = 0
     JohnnyCajeD = 0
     ReptileD = 0
     BarakaD = 0
     ShangTsungD = 0
     JadeD = 0
     DefeatedAll = False
     CharactersResetLight
     LightEffect 1
     BackFlashEffect 3
     BackFlashEffect 4
End Sub

Sub CharactersResetLight()
    Light1.state = 1: If B2SOn Then Controller.B2SSetData 1, 1
    Light2.state = 1: If B2SOn Then Controller.B2SSetData 2, 1
    Light3.state = 1: If B2SOn Then Controller.B2SSetData 3, 1
    Light4.state = 1: If B2SOn Then Controller.B2SSetData 4, 1
    Light5.state = 1: If B2SOn Then Controller.B2SSetData 5, 1
    Light6.state = 1: If B2SOn Then Controller.B2SSetData 6, 1
    Light7.state = 1: If B2SOn Then Controller.B2SSetData 7, 1
    Light8.state = 1: If B2SOn Then Controller.B2SSetData 8, 1
    Light9.state = 1: If B2SOn Then Controller.B2SSetData 9, 1
    Light10.state = 1: If B2SOn Then Controller.B2SSetData 10, 1
    Light11.state = 1: If B2SOn Then Controller.B2SSetData 11, 1
    Light12.state = 1: If B2SOn Then Controller.B2SSetData 12, 1
    Light13.state = 1: If B2SOn Then Controller.B2SSetData 13, 1
    Light14.state = 1: If B2SOn Then Controller.B2SSetData 14, 1
    Light15.state = 1: If B2SOn Then Controller.B2SSetData 15, 1
    Light16.state = 1: If B2SOn Then Controller.B2SSetData 16, 1
End Sub



'*******************
' Targets Multiball
'*******************
Sub MortalKombatTargets 
  If bMultiBallMode = True Then Exit Sub
  if bsLocksTargets1 + bsLocksTargets2 + bsLocksTargets3 = 3 Then
     DMDFlush
     DMD RightLine(1, ""), RightLine(1, ""), 28, eNone, eNone, eBlink, 2500, True, "KahnPrepareToDie"
     DMD_DisplaySceneEx "DMD1.png", "Targets", 14, 2, "Multiball" , -1, -1, UltraDMD_Animation_None, 2000, UltraDMD_Animation_None
     ChangeGi "amber"
     AddMultiball 3
     EnableBallSaver 10 
     LockLight.State = 0
    'Turn On the Jackpot lights
     LJ1.state = 2
     LJ2.state = 2
     LJ3.state = 2
     LJ4.state = 2
     CenterRampLight.state = 0
     RightRampLight.state = 0
     SetFlash 3, 0
     SetFlash 4, 0
     SetFlash 5, 0
     MKIC = 0
     MKIIC = 0
     MKIIIC = 0
     bsLocksTargets1 = 0
     bsLocksTargets2 = 0
     bsLocksTargets3 = 0
	 If Sw25.Isdropped = 0 Then DOF 114, DOFPulse
     Sw25.isdropped = 1
     l18.State = 0
     Post.Isdropped= 0
     LightEffect 1
  End If
end Sub

'*****************
' Locks Multiball
'*****************
Sub MKMultiball() 
    EnableBallSaver 10
  '  ChangeSong
     DMDFlush
    DMD RightLine(1, ""), RightLine(1, ""), 29, eNone, eNone, eBlink, 2500, True, "KahnPrepareToDie"
     DMD_DisplaySceneEx "DMD1.png", "", 14, 2, "   Multiball   " , -1, -1, UltraDMD_Animation_None, 2000, UltraDMD_Animation_None
    AddMultiball 2 
    ChangeGi "amber"
    ' turn off the lock lights
    LockedBalls = 0
    bLockEnabled = False
    LockLight.State = 0
	If Sw25.Isdropped = 0 Then DOF 114, DOFPulse
    Sw25.isdropped = 1
    l18.State = 2
    Post.Isdropped = 0
    'Turn On the Jackpot lights
     LJ1.state = 2
     LJ2.state = 2
     LJ3.state = 2
     LJ4.state = 2
    CenterRampLight.state = 0
    RightRampLight.state = 0
    BackFlashEffect 3
    BackFlashEffect 4
    LightEffect 1
End Sub

'***********************
' Craracters Multiballs
'***********************

Dim KinMultiball: KinMultiball = 0 
Sub KintaroMultiball()
  If bMultiBallMode = True Then Exit Sub
  if KinMultiball = 1 Then
     DMDFlush
     DMD RightLine(1, ""), RightLine(1, ""), 36, eNone, eNone, eBlink, 2500, True, "KahnPrepareToDie"
     DMD_DisplaySceneEx "DMD1.png", "Kintaro", 14, 2, "Multiball" , -1, -1, UltraDMD_Animation_None, 2000, UltraDMD_Animation_None
     ChangeGi "orange"
     AddMultiball 2
     EnableBallSaver 10 
     LockLight.State = 0
    'Turn On the Jackpot lights
     LJ1.state = 2
     LJ2.state = 2
     LJ3.state = 2
     LJ4.state = 2
     CenterRampLight.state = 0
     RightRampLight.state = 0
     SetFlash 3, 0
     SetFlash 4, 0
     SetFlash 5, 0
	 If Sw25.Isdropped = 0 Then DOF 114, DOFPulse
     Sw25.isdropped = 1
     l18.State = 0
     Post.Isdropped= 0
     KintaroD = 0
     ShaoKahnD = 1
     CharacterMode = True
     LightEffect 1
     BackFlashEffect 3
     BackFlashEffect 4
  End If
End Sub

Dim UMultiball: UMultiball = 0
Sub UnknownMultiball()
  If bMultiBallMode = True Then Exit Sub
  if UMultiball = 1 Or UMultiball = 2 Then
    If UMultiball = 1 Then ChangeGi "purple" End If
    If UMultiball = 2 Then ChangeGi "green" End If
     DMDFlush
     DMD RightLine(1, ""), RightLine(1, ""), 35, eNone, eNone, eBlink, 2500, True, "KahnPrepareToDie"
     DMD_DisplaySceneEx "DMD1.png", "?", 15, 6, "Multiball" , -1, -1, UltraDMD_Animation_None, 2000, UltraDMD_Animation_None
     AddMultiball 1
     EnableBallSaver 10 
     LockLight.State = 0
    'Turn On the Jackpot lights
     LJ1.state = 2
     LJ2.state = 2
     LJ3.state = 2
     LJ4.state = 2
     CenterRampLight.state = 0
     RightRampLight.state = 0
     SetFlash 3, 0
     SetFlash 4, 0
     SetFlash 5, 0
	 If Sw25.Isdropped = 0 Then DOF 114, DOFPulse
     Sw25.isdropped = 1
     l18.State = 0
     Post.Isdropped= 0
     CharacterMode = True
     LightEffect 1
     BackFlashEffect 3
     BackFlashEffect 4
  End If
End Sub

Sub CharactersDefeated()
 If CharacterMode = True Then
   If KinMultiball = 1 Then
          if CenterRampLight.state = 2 Or RightRampLight.state = 2 Then
             RightRampLight.State = 0
             CenterRampLight.State = 0
             JackpotKintaroDefeated()       
          End If
   End If

   If UMultiball = 1 Then
          if CenterRampLight.state = 2 Or RightRampLight.state = 2 Then
             RightRampLight.State = 0
             CenterRampLight.State = 0
             JackpotUnknownCharacterDefeated()   
          End If
   End If
 End If
End Sub

Sub StopAllMusic:EndMusic:End Sub
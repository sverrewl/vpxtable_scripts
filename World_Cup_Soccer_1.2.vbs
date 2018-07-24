Option Explicit

' Thalamus 2018-07-24
' Table doesn't have "Positional Sound Playback Functions" or "Supporting Ball & Sound Functions"
' No special SSF tweaks yet.


Dim DesktopMode: DesktopMode = tablewpc94.ShowDT
Dim KDBall : KDBall = True
Scoretext.visible = DesktopMode

dim UseVPMDMD:UseVPMDMD = DesktopMode
dim UseVPMModSol:UseVPMModSol = 1
dim BrightFlashers: BrightFlashers = True

Dim cController, ROL, Hidden, DefaultOptions
DefaultOptions = 1*optController + 2*optB2BEnable + 2*optGoalieSpeed

Const cGameName = "wcs_l2"

Dim FeedbackSounds:FeedbackSounds = Array("ballrel","bumper","diverter","flipperup","flipperdown","knocker","popper","popper_ball","solenoid","target","lsling","solon","jet3")
'*** End Options ***

LoadVPM "01530000", "WPC.VBS", 3.10


 Sub LoadVPM(VPMver, VBSfile, VBSver)   'Add new call to InitializeOptions to allow selection of controller through F6 menu
    On Error Resume Next
        If ScriptEngineMajorVersion < 5 Then MsgBox "VB Script Engine 5.0 or higher required"
            ExecuteGlobal GetTextFile(VBSfile)
        If Err Then MsgBox "Unable to open " & VBSfile & ". Ensure that it is in the same folder as this table. " & vbNewLine & Err.Description
        If VPinMAMEDriverVer < VBSver Or Err Then MsgBox VBSFile & " ver " & VBSver & " or higher required."

        InitializeOptions 'Enables New Controller change through F6 menu, so it needs to be placed before Controller selection

        Select Case cController
            Case 1:
                Set Controller = CreateObject("VPinMAME.Controller")
                If Err Then MsgBox "Can't Load VPinMAME." & vbNewLine & Err.Description
                If VPMver>"" Then If Controller.Version < VPMver Or Err Then MsgBox "VPinMAME ver " & VPMver & " required."
            Case 2:
                Set Controller = CreateObject("UltraVP.BackglassServ")
            Case 3:
                Set Controller = CreateObject("B2S.Server")
        End Select
        If Err then
            msgbox "Invalid controller selected, defaulting to VPinMame"
            Set controller = CreateObject("VPinMAME.Controller")
        End If
    On Error Goto 0
 End Sub

Const UseSolenoids = True
Const UseLamps     = False
Const UseSync      = True
Const UseGI        = True
'******GI CALL********'
 Dim LeftGIs, RightGIs, UpGIs
 Dim GILevels
' InitGI
'Set GiCallback2 = GetRef("UpdateGI2")
Set GICallback = GetRef("UpdateGI")
'Set MotorCallback = GetRef("UpdateFlipperLogos")
' Standard Sounds used by Driver help code
'Const SSolenoidOn   = "SolOn"
'Const SSolenoidOff  = "SolOff"
Const SFlipperOn    = "FlipperUp"
Const SFlipperOff   = "FlipperDown"
Const SCoin="Coin"

'--------------------------------
' Init the table, Start VPinMAME
'--------------------------------
Dim bsTrough, bsLock, bsLeftEject, bsRightEject, bsLeft, bsRight, bsUpper, bsGoal, bsVUK, bsTV
Dim ttBall,mGoalie,mBall,vlLock, LockMagnetSave, MagnaGoalie, plungerIM, x
Dim BallRel, SolOn, Popper
Dim RefreshARlight:RefreshARlight=False
Sub TableWPC94_Init
    vpmInit me
    On Error Resume Next
        With Controller
            .GameName = cGameName
            .Games(cGameName).Settings.Value("rol") = ROL   'Set through the F6 menu
            .Hidden = DesktopMode                               'Set through the F6 menu
            .SplashInfoLine = "World Cup Soccer 94, Bally 1994"
            .HandleMechanics = 0
            .HandleKeyboard = False
            .ShowDMDOnly = True : .ShowFrame = False : .ShowTitle = False
            .Run : If Err Then MsgBox Err.Description : Exit Sub
        End With
    On Error Goto 0 'Create Controller Object, and read in options

    if DesktopMode Then
        ' Ugh
        F68.x = F68.x - 10
        F68.y = F68.y + 40
        F86.x = F85.x - 15
        F86.y = F86.y + 50
        F85.x = F85.x - 20
        F85.y = F85.y + 50

		F71A.height = 260.0
		F76.height = 170
		F78A.height = 215

		F77.height = 140
		F78.height = 180
		F71.height = 225

    end if

    vpmNudge.TiltSwitch = swTilt
    vpmNudge.Sensitivity = 5

    ' Main Timer init
    PinMAMETimer.Interval = PinMAMEInterval
    PinMAMETimer.Enabled = true

    'Impulse Plunger
    '--------------------
    Set plungerIM = New cvpmImpulseP
    With plungerIM
        .InitImpulseP swPlunger, IMPowerSetting, IMTime
        .Random IMScatter
        .InitEntrySnd "PlungerPull"
        .InitExitSnd "plunger2", "plunger"
        .CreateEvents "plungerIM"
    End With

'------------------------------
' Set Up Ballstacks and init info
'------------------------------
    Set bsTrough = new cvpmBallStack
        bsTrough.InitSw 0,31,32,33,34,35,0,0
        bsTrough.InitKick BallRelease,40,8
        bsTrough.Balls = 5
        bsTrough.InitExitSnd BallRel,SolOn

    Set bsLeft = New cvpmBallStack
        bsLeft.InitSaucer LeftEjectHole, swLeftEjectHole,150,7
        bsLeft.InitExitSnd BallRel,SolOn

    Set bsRight = New cvpmBallStack
        bsRight.InitSaucer RightEjectHole, swRightEjectHole,210,8
        bsRight.InitExitSnd BallRel,SolOn

    Set bsUpper = New cvpmBallStack
        bsUpper.InitSaucer UpperEjectHole, swUpperEjectHole,10,18
'       bsupper.kickZ = 3.1415926/4
        bsupper.InitExitSnd BallRel,SolOn

    Set bsTV = New cvpmBallStack
        bsTV.InitSw 0,swTVBallPopper,0,0,0,0,0,0
        bsTV.InitKick TVBallPopper, 213, 10
        bsTV.KickBalls = 2
        bsTV.InitExitSnd Popper,SolOn

    Set bsGoal = New cvpmBallStack
        bsGoal.InitSw 0,swGoalTrough,0,0,0,0,0,0
        bsGoal.InitExitSnd Popper,SolOn

    Set bsVUK = New cvpmBallStack
        bsVUK.InitSw 0,swGoalPopperOpto,0,0,0,0,0,0
        bsVUK.InitKick VUKKicker,146, 30  'ORIGINALE 146, 13
        bsVUK.KickZ = Pi/2
        bsVUK.InitExitSnd Popper,SolOn


'Setup magnets
    Set LockMagnetSave = New cvpmMagnet : With LockMagnetSave
        .InitMagnet LockMagnet, 25
        .Solenoid = sLockMagnet
        .CreateEvents "LockMagnetSave"
    End With
    Set MagnaGoalie = New cvpmMagnet : With MagnaGoalie
        .InitMagnet TrgMagnaGoalie, 15 'pretty weak orig at 15 testing
        .Solenoid = sMagnaGoalie
        .GrabCenter = True
        .CreateEvents "MagnaGoalie"
    End With

    ' Visible Lock
    Set vlLock = New cvpmVLock : With vlLock
        .InitVLock Array(LockMechLow, LockMechHigh), Array(LockKickLow, LockKickHigh), Array(swLockMechLow, swLockMechHigh)
        .CreateEvents "vlLock"
    End With

    ' Spinning ball
    Set ttBall = New cvpmturnTable : With ttBall
        .InitTurnTable BallTrigger, 100
        .SpinUp = 0 : .SpinDown = 0
        .CreateEvents "ttBall"
    End With

    ' Mechs
    Set mGoalie = new cvpmMech : With mGoalie
        .Sol1 = sGoalieDrive
        .MType = vpmMechLinear + vpmMechReverse + vpmMechOneSol' + vpmMechFast
        .Length = GoalieSpeed
        .Steps = 320
        .AddSw swGoalIsLeft, 0, 8
        .AddSw swGoalIsRight, 152,168
        .AddSw swGoalIsLeft, 312, 320
        .Callback = GetRef("DrawGoalie")
        .Start
    End With

	if not KDBall then
		Set mBall = new cvpmMech : With mBall
			.Sol1 = sBallClockwise : .Sol2 = sBallCounterCW
			.MType = vpmMechLinear + vpmMechCircle + vpmMechTwoDirSol
			.Acc = 60 : .Ret = 2
			.Length = 12
			.Steps = 24
			.Callback = GetRef("UpdateBall")
			.Start
		End With
	end if

    Set MotorCallback = GetRef("UpdateFlipperLogos")

    GWall0.isdropped=1
    GWall1.isdropped=1
    GWall2.isdropped=1
    GWall3.isdropped=1
    GWall4.isdropped=1
    GWall5.isdropped=1
    GWall6.isdropped=1
    GWall7.isdropped=1
    GWall8.isdropped=1
    GWall9.isdropped=1
    GWall10.isdropped=1
    GWall11.isdropped=1
    GWall12.isdropped=1
    GWall13.isdropped=1
    GWall14.isdropped=1
    GWall15.isdropped=1
    GWall16.isdropped=1
    GWall17.isdropped=1
    GWall18.isdropped=1
    GWall19.isdropped=1
    GWall20.isdropped=1
    Controller.Switch(swCoinDoor) = 1
End Sub

'-------------------
' keyboard routines
'-------------------
'ExtraKeyHelp = KeyName(keyFront) & vbTab & "Buy-in Button" & vbNewLine &_
'               KeyName(keyUpperLeft) & vbTab & "Magna Goalie"

Sub TableWPC94_KeyUp(ByVal keycode)
        If keycode = PlungerKey Then
        Plunger.Fire
        StopSound "PlungerPull"
        PlaySound "plunger3"
    End If
    If (keycode = rightmagnasave or KeyCode = 3) Then Controller.Switch(swBuyInButton) = False
    If KeyCode = leftmagnasave Then Controller.Switch(swMagGoalieButton) = False
    If KeyUpHandler(keycode) Then Exit Sub
End Sub

Sub TableWPC94_KeyDown(ByVal keycode)
    If keycode = PlungerKey Then
        Plunger.PullBack
        PlaySound "PlungerPull"
    End If
    If (keycode = rightmagnasave or KeyCode = 3) Then Controller.Switch(swBuyInButton) = True
    If KeyCode = leftmagnasave Then Controller.Switch(swMagGoalieButton) = True
    If KeyDownHandler(keycode) Then Exit Sub
End Sub

'        Koadic's Alpha Ramp
'   Impulse Plunger Scripting v6
'       single ramp animated
'        via image switching
'------------------------------
 Dim PDelay, PCount, PImages, PStart, IMTime, IMPowerSetting, PlFrame, IMScatter

 IMPowerSetting = Plunger.MechStrength  ' Plunger Power - Set via Plunger MechStrength
 IMTime = Round(Plunger.PullSpeed/10, 2)' Time in 1/10th seconds for Full Plunge - Set via Plunger Pull Speed...
                                            ' 1 = .1 second, 5 = .5 second, 10 = 1 second, etc.
 IMScatter = Plunger.ScatterVelocity    ' Plunger Scatter Velocity - Percentage of variation in Plunger Power
    ' Setting Scatter Velocity to 10 = 10%, if Power is 50, max plunge will vary from 47.5 to 52.5 (+/- 5%)
 PStart = 0             ' Set number of first plunger image, use 1 for legacy "1-12" setup
 PImages = 25           ' Set number of animation frames not including the PStart position, use 11 for legacy "1-12" setup
 PTime.Interval = INT(IMTime*1000/PImages)

 PDelay = CINT(Plunger.FireSpeed/Plunger.TimerInterval)
 ReDim PlPos(PDelay)

 Sub PTime2_Timer
    Select Case PCount
        Case 0:aPlunger.Image = "p" & PStart : PRefresh.state = ABS(PRefresh.state - 1)
        Case 1:aPlunger.Image = "p" & INT(PImages/5) : PRefresh.state = ABS(PRefresh.state - 1)
        Case 2:ResetPlungers:Plunger.TimerEnabled = 1:Me.Enabled = 0
    End Select
    Pcount = Pcount + 1
 End Sub

 Sub Plunger_Timer()
    PlPos(PDelay) = Plunger.Position
    PlFrame = PlPos(PDelay)
    If PlPos(PDelay) <> PlPos(PDelay - 1) Then
        aPlunger.Image = "p" & PlFrame
        PRefresh.state = ABS(PRefresh.state - 1)
        If PlPos(PDelay) < 2 and PlPos(0) > 5 Then
            PlungerIM.Strength = (PlPos(0)/25*Plunger.MechStrength)
            PlungerIM.AutoFire
            PlungerIM.Strength = Plunger.MechStrength
            Plunger.TimerEnabled = 0:PTime2.Enabled = 1
        End If
    End If
    For x = 0 to ubound(PlPos)-1:PlPos(x)=PlPos(x+1):Next
 End Sub

'----------------
' Goalie Mech
'----------------
Dim GoalieWalls, GIWalls
GoalieWalls = Array(GWall0, GWall1, GWall2, GWall3, GWall4, GWall5, GWall6, GWall7, GWall8, GWall9, GWall10, GWall11, GWall12, GWall13, GWall14, GWall15, GWall16, GWall17, GWall18, GWall19, GWall20)



Sub DrawGoalie(aCurrPos,aSpeed,aLast)
    GoalieWalls(Int(160-ABS(aLast-160))/8).IsDropped = True
    GoalieWalls(INT(160-ABS(aCurrPos-160))/8).IsDropped = False
    Goalie.roty = dSin((80 - ABS(160-aCurrPos)) * (9/8)) * 10
End Sub

Sub UpdateBall(aCurrPos,aSpeed,aLast)
	if not KDBall then
		ttBall.MotorOn = aSpeed <> 0
		ttBall.Speed = aSpeed
		SoccerBall.rotz = SoccerBall.rotz + aSpeed/1.5
	end if
    'SoccerBall.TriggerSingleUpdate
End Sub


Const SoccerMaxSpeed = 200
Const SoccerMaxVisibleSpeed = 30 ' After 30 the ball appears to reverse.
Const SoccerAccelMotorOn = 3.0
Const SoccerAccelMotorOff = 2.5

Dim SoccerTargetSpeed, SoccerAccel

'--------------------------------------------
' Spinning Ball Handler (KieferSkunk/Dorsola)
'--------------------------------------------

' Instead of using a cvpmMech, which can't correctly handle both solenoids
' being on at the same time, we're using a more "manual" method of controlling
' the turntable, and thus the spinning ball animation.  This gives us greater
' control over the solenoid logic.
'
' I spent a lot of time studying the solenoid patterns in VPinMAME, watching
' YouTube videos of the soccer ball during various scenarios in the game, and
' studying the DC motor control board schematic and that of its L6203 chip.
' VPinMAME reports only four binary states for the two solenoids:
'
' CW   CCW  Result
' OFF  OFF  Ball stopped
' ON   ON   Ball spins forward at half speed (eg. between jackpots)
' ON   OFF  Ball spins forward at full speed (goal lit)
' OFF  ON   Ball spins backward at full speed (Final Match)
'
' I'm still not convinced that this is actually how it works, but this version
' of the sim is better than what I had before.  However, the ball still spins
' in the wrong direction for the first few seconds of Final Match, unlike the
' real machine, so either VPinMAME is emulating the solenoids incorrectly, or
' there's more going on in the electronics (eg. high-speed pulsing) that the
' emu isn't telling us.  Until we track down what's going on in the ROM itself
' and/or the real electronics, this will probably have to do.

Dim SoccerCWOn, SoccerCCWOn
SoccerCWOn = False
SoccerCCWOn = False

Sub solBallClockwise(Enabled)
    SoccerCWOn = Enabled
    ModifyTurntableState
End Sub

Sub SolBallCounterCW(Enabled)
    SoccerCCWOn = Enabled
    ModifyTurntableState
End Sub

Sub ModifyTurntableState()
    If SoccerCWOn AND SoccerCCWOn Then
        ' Both solenoids on - ball spins forward at half speed
        SoccerTargetSpeed = SoccerMaxSpeed * 0.5
        SoccerAccel = SoccerAccelMotorOn
    ElseIf SoccerCWOn Then
        ' Full-speed clockwise
        SoccerTargetSpeed = SoccerMaxSpeed
        SoccerAccel = SoccerAccelMotorOn
    ElseIf SoccerCCWOn Then
        ' Full-speed counter-clockwise
        SoccerTargetSpeed = -SoccerMaxSpeed
        SoccerAccel = SoccerAccelMotorOn
    Else
        ' Motor off.
        SoccerTargetSpeed = 0
        SoccerAccel = SoccerAccelMotorOff
    End If
End Sub

Sub AnimateBallTimer_Timer()
    ' Adjust ball's speed toward target, if necessary
    If ttBall.Speed > SoccerTargetSpeed Then
        ttBall.Speed = Max(ttBall.Speed - SoccerAccel, SoccerTargetSpeed)
    ElseIf ttBall.Speed < SoccerTargetSpeed Then
        ttBall.Speed = Min(ttBall.Speed + SoccerAccel, SoccerTargetSpeed)
    End If

    ' Animate ball if it's moving.
    If ttBall.Speed Then
        ' Adjust primitive's absolute rotation amount by our current speed.
        ' (Note: TurnTable ball-effect factor is double animation speed)
         SoccerBall.RotZ = SoccerBall.RotZ + ttBall.Speed * (SoccerMaxVisibleSpeed / SoccerMaxSpeed)
        If SoccerBall.RotZ > 360 Then SoccerBall.RotZ = SoccerBall.RotZ - 360
        If SoccerBall.RotZ < -360 Then SoccerBall.RotZ = SoccerBall.RotZ + 360
	' Experimental blur code.  Reallly need re-textured ball for this.
	'	if ttBall.Speed > SoccerMaxSpeed * 0.5 Then
	'		SoccerBall.Material = "SoccerBlur1"
	'		SoccerBallBlur.RotZ = SoccerBall.RotZ - 30
	'		SoccerBallBlur.Visible =true
	'	Else
	'		SoccerBall.Material = "Plastic with an image"
	'		SoccerBallBlur.Visible =false
	'	end if
	End If
End Sub

Function Min(value1, value2)
    If value1 < value2 Then
        Min = value1
    Else
        Min = value2
    End If
End Function

Function Max(value1, value2)
    If value1 > value2 Then
        Max = value1
    Else
        Max = value2
    End If
End Function

'--------------------------
' Goal & VUK handling
'--------------------------
Sub HandleGoalTrough(swNo)
    bsGoal.AddBall 0
    If bsVUK.Balls = 0 Then vpmTimer.AddTimer 100, "ExitGoal"
End Sub

Sub ExitGoal(swNo)
    If bsVUK.Balls = 0 And bsGoal.Balls > 0 Then
        bsGoal.SolOut True : bsVUK.AddBall 0
    End If
End Sub

Sub SolVUK(aEnabled)
    if aEnabled Then bsVUK.SolOut True : ExitGoal 0
End Sub

'------------------------
' Lock
'------------------------
Sub MagnaLock_Hit : Me.Enabled = False : End Sub
Sub SolMagnaLock(aEnabled)
    MagnaLock.Enabled = aEnabled
    If Not aEnabled Then MagnaLock.Kick 195,1
End Sub

Sub Solenoide_Hit
    vpPlay "MagneteL",ActiveBall
End Sub

'----------------------------
' Kicker Switches
'----------------------------
Sub UpperEjectHole_Hit   : bsUpper.AddBall Me    : vpPlay "EnterHoleU", ActiveBall : End Sub
Sub UpperEjectHole_UnHit : vpPlay "ExitKicher", ActiveBall: End Sub
Sub RightEjectHole_Hit   : bsRight.AddBall Me    : vpPlay "EnterHoleR", ActiveBall : End Sub
Sub RightEjectHole_UnHit : vpPlay "ExitKicher", ActiveBall: End Sub
Sub LeftEjectHole_Hit    : bsLeft.AddBall Me     : vpPlay "EnterHoleL", ActiveBall : End Sub
Sub LeftEjectHole_UnHit  : vpPlay "ExitKicher", ActiveBall: End Sub
Sub TVBallPopper_Hit     : StopRollingSound: ClearBallid : bsTV.AddBall Me       : vpPlay "kicker_enter", TVBallPopper : End Sub
Sub GoalPopperOpto_Hit   : StopRollingSound: ClearBallid :bsVUK.AddBall Me       : vpPlay "kicker_enter", GoalPopperOpto : End Sub
Sub Drain_Hit            : StopRollingSound: ClearBallid : bsTrough.AddBall Me   : StopRollingSound: vpPlay "Drain",Drain : Playsound "drain":End Sub
Sub GoalTrough1_Hit      : vpPlay "GoalEnter", ActiveBall : ClearBallid : Me.DestroyBall : vpmTimer.AddTimer 110,"HandleGoalTrough" : End Sub
Sub GoalTrough2_Hit      : vpPlay "GoalEnter", ActiveBall : ClearBallid : Me.DestroyBall : vpmTimer.AddTimer 90,"HandleGoalTrough" : End Sub
Sub GoalTrough3_Hit      : vpPlay "GoalEnter", ActiveBall : ClearBallid : Me.DestroyBall : vpmTimer.AddTimer 70,"HandleGoalTrough" : End Sub
Sub GoalTrough4_Hit      : vpPlay "GoalEnter", ActiveBall : ClearBallid : Me.DestroyBall : vpmTimer.AddTimer 50,"HandleGoalTrough" : End Sub
Sub GoalTrough5_Hit      : vpPlay "GoalEnter", ActiveBall : ClearBallid : Me.DestroyBall : vpmTimer.AddTimer 30,"HandleGoalTrough" : End Sub
Sub GoalTrough6_Hit      : vpPlay "GoalEnter", ActiveBall : ClearBallid : Me.DestroyBall : vpmTimer.AddTimer 10,"HandleGoalTrough" : End Sub
'-----------------------------------
'Switch Routines
'-----------------------------------
Sub BallShooter_Hit         : Controller.Switch(swBallShooter) = true           : End Sub
Sub BallShooter_Unhit       : Controller.Switch(swBallShooter) = false          : End Sub
Sub SkillShotFront_Hit      : vpPlay "DropRampSkillShot", SkillShotFront :vpmTimer.PulseSw swSkillShotFront                 : End Sub
Sub SkillShotCenter_Hit     : vpPlay "DropRampSkillShot", SkillShotCenter :vpmTimer.PulseSw swSkillShotCenter               : End Sub
Sub SkillShotRear_Hit       : vpPlay "DropRampSkillShot", SkillShotRear :vpmTimer.PulseSw swSkillShotRear                   : End Sub


Sub RightOutLane_Hit() ' Kickback
    'RightOutLane_a.IsDropped = 0
    Controller.Switch(swRightOutLane) = 1
    vpPlay "sensor", RightOutLane
End Sub
Sub RightOutLane_Unhit()
    'RightOutLane_a.IsDropped = 1
    Controller.Switch(swRightOutLane) = 0
End Sub

Sub RightFlipperLane_Hit() ' Kickback
    'RightFlipperLane_a.IsDropped = 0
    Controller.Switch(swRightFlipperLane) = 1
    vpPlay "sensor", RightFlipperLane
End Sub
Sub RightFlipperLane_Unhit()
    'RightFlipperLane_a.IsDropped = 1
    Controller.Switch(swRightFlipperLane) = 0
End Sub

Sub LeftFlipperLane_Hit() ' Kickback
    'LeftFlipperLane_a.IsDropped = 0
    Controller.Switch(swLeftFlipperLane) = 1
    vpPlay "sensor", LeftFlipperLane
End Sub
Sub LeftFlipperLane_Unhit()
    'LeftFlipperLane_a.IsDropped = 1
    Controller.Switch(swLeftFlipperLane) = 0
End Sub

Sub Kickback_Hit() ' Kickback
    'Kickback_a.IsDropped = 0
    Controller.Switch(swKickback) = 1
    vpPlay "sensor", Kickback
End Sub
Sub Kickback_UnHit()
    'Kickback_a.IsDropped = 1
    Controller.Switch(swKickback) = 0
End Sub


'Sub Kickback_Hit           : Controller.Switch(swKickback) = true          : End Sub
'Sub Kickback_Unhit         : Controller.Switch(swKickback) = false         : End Sub
'Sub LightMagGoalie_Hit:LightMagGoalie.IsDropped = TRUE:LightMagGoaliea.IsDropped = FALSE:Me.TimerEnabled = 1:vpmTimer.PulseSw (swLightMagGoalie):vpPlay "target":End Sub
'Sub LightMagGoalie_Timer:LightMagGoalie.IsDropped = FALSE:LightMagGoaliea.IsDropped = TRUE:Me.TimerEnabled = 0:End Sub
Sub LightMagGoalie_Hit      : vpmTimer.PulseSw swLightMagGoalie                 : End Sub
'Sub LightKickback_Hit:LightKickback.IsDropped = TRUE:LightKickbacka.IsDropped = FALSE:Me.TimerEnabled = 1:vpmTimer.PulseSw (swLightKickback):vpPlay "target":End Sub
'Sub LightKickback_Timer:LightKickback.IsDropped = FALSE:LightKickbacka.IsDropped = TRUE:Me.TimerEnabled = 0:End Sub
Sub LightKickback_Hit       : vpmTimer.PulseSw swLightKickback                  : End Sub
Sub Spinner1_Spin           : vpmTimer.PulseSw swSpinner    : vpPlay "spinner", Spinner1                    : End Sub
Sub LeftSlingshot_Slingshot : vpmTimer.PulseSw swLeftSlingshot:vpPlay "SlingshotSinistro", Sling1   : End Sub
Sub RightSlingshot_Slingshot: vpmTimer.PulseSw swRightSlingshot:vpPlay "SlingshotDestro", Sling2        : End Sub
Sub LeftJetBumper_Hit       : vpmTimer.PulseSw swLeftJetBumper :vpPlay "BumperSinistro", LeftJetBumper    : End Sub
Sub UpperJetBumper_Hit      : vpmTimer.PulseSw swUpperJetBumper:vpPlay "BumperDestro", UpperJetBumper       : End Sub
Sub LowerJetBumper_Hit      : vpmTimer.PulseSw swLowerJetBumper:vpPlay "BumperCentrale", LowerJetBumper     : End Sub
'Sub UpperLeftLane_Hit      : Controller.Switch(swUpperLeftLane) = true          : End Sub
'Sub UpperLeftLane_Unhit    : Controller.Switch(swUpperLeftLane) = false         : End Sub
'Sub UpperRightLane_Hit     : Controller.Switch(swUpperRightLane) = true         : End Sub
'Sub UpperRightLane_Unhit   : Controller.Switch(swUpperRightLane) = false        : End Sub
'Sub FreeKickTarget_Hit:FreeKickTarget.IsDropped = TRUE:FreeKickTargeta.IsDropped = FALSE:Me.TimerEnabled = 1:vpmTimer.PulseSw (swFreeKickTarget):vpPlay "target":End Sub
'Sub FreeKickTarget_Timer:FreeKickTarget.IsDropped = FALSE:FreeKickTargeta.IsDropped = TRUE:Me.TimerEnabled = 0:End Sub
Sub FreeKickTarget_Hit      : vpmTimer.PulseSw swFreeKickTarget                 : End Sub
'Sub KickbackUpper_Hit      : vpmTimer.PulseSw swKickbackUpper                  : End Sub
Sub UpperLeftLane_Hit() ' Kickback
    'UpperLeftLane_a.IsDropped = 0
    Controller.Switch(swUpperLeftLane) = 1
    vpPlay "sensor", UpperLeftLane
End Sub

'************************************************
'************Slingshots Animation****************
'************************************************

Dim RStep, Lstep

'Sub LeftSlingShot_Slingshot: vpmTimer.PulseSw 26: End Sub
'Sub RightSlingShot_Slingshot: vpmTimer.PulseSw 27: End Sub

Sub solLSling(enabled)
    If enabled then
        'PlaySound SoundFX ("SlingshotSinistro",DOFContactors)
        LSling.Visible = 0
        LSling1.Visible = 1
        sling1.TransZ = -27
        LStep = 0
        LeftSlingShot.TimerEnabled = 1
    End If
End Sub

Sub solRSling(enabled)
    If enabled then
        'PlaySound SoundFX ("SlingshotDestro",DOFContactors)
        RSling.Visible = 0
        RSling1.Visible = 1
        sling2.TransZ = -27
        RStep = 0
        RightSlingShot.TimerEnabled = 1
    End If
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling1.TransZ = -15
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling1.TransZ = 0:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling2.TransZ = -15
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling2.TransZ = 0:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub








Sub UpperLeftLane_UnHit()
    'UpperLeftLane_a.IsDropped = 1
    Controller.Switch(swUpperLeftLane) = 0
End Sub

Sub UpperRightLane_Hit() ' Kickback
    'UpperRightLane_a.IsDropped = 0
    Controller.Switch(swUpperRightLane) = 1
    vpPlay "sensor", UpperRightLane
End Sub
Sub UpperRightLane_UnHit()
    'UpperRightLane_a.IsDropped = 1
    Controller.Switch(swUpperRightLane) = 0
End Sub

Sub KickbackUpper_Hit() ' Kickback
    'KickbackUpper_a.IsDropped = 0
    Controller.Switch(swKickbackUpper) = 1
    vpPlay "sensor", KickbackUpper
End Sub
Sub KickbackUpper_UnHit()
    'KickbackUpper_a.IsDropped = 1
    Controller.Switch(swKickbackUpper) = 0
End Sub


Sub Rollover1_Hit           : Controller.Switch(swRollover1) = true             : End Sub
Sub Rollover1_Unhit         : Controller.Switch(swRollover1) = false            : End Sub
Sub Rollover2_Hit           : Controller.Switch(swRollover2) = true             : End Sub
Sub Rollover2_Unhit         : Controller.Switch(swRollover2) = false            : End Sub
Sub Rollover3_Hit           : Controller.Switch(swRollover3) = true             : End Sub
Sub Rollover3_Unhit         : Controller.Switch(swRollover3) = false            : End Sub
Sub Rollover4_Hit           : Controller.Switch(swRollover4) = true             : End Sub
Sub Rollover4_Unhit         : Controller.Switch(swRollover4) = false            : End Sub
Sub Striker1_Hit            : vpmTimer.PulseSw swStriker1 : vpPlay "target", ActiveBall : End Sub
Sub Striker2_Hit            : vpmTimer.PulseSw swStriker2: vpPlay "target", ActiveBall : End Sub
Sub Striker3High_Hit        : vpmTimer.PulseSw swStriker3High : vpPlay "target", ActiveBall  : End Sub
Sub LeftRampEntrance_Hit    : Controller.Switch(swLeftRampEntrance) = true      : vpPlay "gate", LeftRampEntrance : End Sub
Sub LeftRampEntrance_Unhit  : Controller.Switch(swLeftRampEntrance) = false     : RampaSinistra : End Sub
Sub LeftRampExit_Hit        : Controller.Switch(swLeftRampExit) = true          : End Sub
Sub LeftRampExit_Unhit      : Controller.Switch(swLeftRampExit) = false         : End Sub
Sub RightRampEntrance_Hit   : Controller.Switch(swRightRampEntrance) = true     : vpPlay "gate", RightRampEntrance : End Sub
Sub RightRampEntrance_Unhit : Controller.Switch(swRightRampEntrance) = false    : RampaDestra : End Sub
Sub RightRampExit_Hit       : Controller.Switch(swRightRampExit) = true         : End Sub
Sub RightRampExit_Unhit     : Controller.Switch(swRightRampExit) = false        : End Sub

'Sub TravelLaneRolo_Hit         : Controller.Switch(swTravelLaneRolo) = true        : End Sub
'Sub TravelLaneRolo_Unhit   : Controller.Switch(swTravelLaneRolo) = false       : End Sub

Sub TravelLaneRolo_Hit() ' Kickback
    'TravelLaneRolo_a.IsDropped = 0
    Controller.Switch(swTravelLaneRolo) = 1
    vpPlay "sensor", TravelLaneRolo
End Sub
Sub TravelLaneRolo_Unhit()
    'TravelLaneRolo_a.IsDropped = 1
    Controller.Switch(swTravelLaneRolo) = 0
End Sub

Sub RampaSinistra
    If ActiveBall.velY < 0  Then
        if tablewpc94.VersionMinor > 3 OR tablewpc94.VersionMajor > 10 Then
            PlaySound "EntrataRampa",0,1,Pan(ActiveBall),0,0,1,0,AudioFade(ActiveBall)
            Playsound "plasticrolling",1,1,Pan(ActiveBall),0,0,1,0,AudioFade(ActiveBall)
        else
            PlaySound "EntrataRampa",0,1,Pan(ActiveBall)
            Playsound "plasticrolling",1,1,Pan(ActiveBall)
        End if
    Else
        StopSound "rrenter"
        StopSound "plasticrolling"
    End If
End Sub

Sub RampaDestra
    If ActiveBall.velY < 0  Then
        if tablewpc94.VersionMinor > 3 OR tablewpc94.VersionMajor > 10 Then
            PlaySound "EntrataRampa",0,1,Pan(ActiveBall),0,0,1,0,AudioFade(ActiveBall)
            Playsound "plasticrolling",1,1,Pan(ActiveBall),0,0,1,0,AudioFade(ActiveBall)
        Else
            PlaySound "EntrataRampa",0,1,Pan(ActiveBall)
            Playsound "plasticrolling",1,1,Pan(ActiveBall)
        End If
    Else
        StopSound "EntrataRampa"
        StopSound "plasticrolling"
    End If
End Sub



Sub TackleSwitch_Hit        : vpmTimer.PulseSw swTackleSwitch                   : End Sub
Sub LeftRampDiverter_Hit    : Controller.Switch(swLeftRampDiverter) = true      : End Sub
Sub LeftRampDiverter_Unhit  : Controller.Switch(swLeftRampDiverter) = false     : End Sub
Sub TroughStack_Hit         : Controller.Switch(swTroughStack) = true           : End Sub
Sub TroughStack_Unhit       : Controller.Switch(swTroughStack) = false          : Playsound "ballrelease":End Sub

Sub GWalls_hit(idx):vpmTimer.PulseSw swGoalieTarget:vpPlay "target", ActiveBall:Goalie.rotx=1.5:GoalieHit.enabled=1:End Sub

Sub GoalieHit_Timer
    If me.uservalue = "" Then me.uservalue = 0
    Select Case me.uservalue
        Case 0 : Goalie.rotx = 1:me.uservalue=me.uservalue+1
        Case 1 : Goalie.rotx = 0.5:me.uservalue=me.uservalue+1
        Case 2 : Goalie.rotx = 0:Me.uservalue = 0:Me.enabled = 0:Exit Sub
    End Select
End Sub

Dim SkillVelY

Sub SkillVelTrig_Hit()
    SkillVelY = ActiveBall.VelY
End Sub

'-----------------------------------
'Map Solenoid Subroutines
'-----------------------------------

SolCallback(12) = "solLSling"
SolCallback(13) = "solRSling"


SolCallback(sLRFlipper)       = "SolRFlipper"
SolCallback(sLLFlipper)       = "SolLFlipper"
SolCallback(sDiverterHold)    = "SolRampDiverter"
SolCallback(sLoopGate)        = "vpmSolGate LoopGate1,False,"
SolCallback(sKickback)        = "SolKickBack"
SolCallback(sLockRelease)     = "SolLockRelease"
'SolCallback(sLockMagnet)     = "SolMagnaLock"
SolCallback(sKnocker)         = "SolKnocker"
SolCallback(sGoalPopper)      = "SolVUK"
SolCallback(sUpperEjectHole)  = "bsUpper.SolOut"
SolCallback(sRightEjectHole)  = "bsRight.SolOut"
SolCallback(sLeftEjectHole)   = "bsLeft.SolOut"
SolCallback(sTrough)          = "bsTrough.SolOut"
SolCallback(sTVPopper)        = "bsTV.SolOut"
SolModCallBack(sLtRampEntrance)  = "SetModLamp 125,"
SolModCallBack(sSpinningBall)    = "SetModLamp 122,"
SolModCallback(sFlipperLanes)    = "SetModLamp 127,"
'SolCallback(sJetBumpers)     = "Sol20"
'SolCallback(sGoal)           = "Sol18"
SolModCallback(sRampRear)        = "SetModLamp 128,"
SolModCallback(sLockArea)        = "SetModLamp 126,"
SolModCallback(sGoalCageTop)     = "SetModLamp 117,"
SolModCallback(sSkillshot)       = "SetModLamp 119,"
SolCallback(sBallClockwise)   = "solBallClockwise"
SolCallback(sBallCounterCW)   = "SolBallCounterCW"

Sub SolKnocker(Enabled)
    If Enabled Then vpPlay "Knocker", L83
End Sub


' Solenoids                     |   Status
'-----------------------------------------
Const sGoalPopper        = 1    'installed
Const sTVPopper          = 2    'installed
Const sKickback          = 3    'installed
Const sLockRelease       = 4    'installed
Const sUpperEjectHole    = 5    'installed
Const sTrough            = 6    'installed
Const sKnocker           = 7    'installed
Const sRampDiverter      = 8    'taken care of with Sol 16
Const sLeftJetBumper     = 9    'handled by VP
Const sUpperJetBumper    = 10   'handled by VP
Const sLowerJetBumper    = 11   'handled by VP
'Const sLeftSlingshot     = 12  'handled by VP
'Const sRightSlingshot    = 13  'handled by VP
Const sRightEjectHole    = 14   'installed
Const sLeftEjectHole     = 15   'installed
Const sDiverterHold      = 16   'installed
Const sGoalCageTop       = 17   ' flasher - installed
Const sGoal              = 18   ' flasher - installed
Const sSkillshot         = 19   ' flasher - installed
Const sJetBumpers        = 20   ' flasher - installed
Const sGoalieDrive       = 21   'installed
Const sSpinningBall      = 22   ' flasher
Const sBallClockwise     = 23   'Spin the Ball
Const sBallCounterCW     = 24   'Spin the Ball
Const sLtRampEntrance    = 25   ' flasher
Const sLockArea          = 26   ' flasher - installed
Const sFlipperLanes      = 27   ' flasher
Const sRampRear          = 28   ' flasher - installed
Const sMagnaGoalie       = 33   'installed
Const sLoopGate          = 34   'installed
Const sLockMagnet        = 35   'installed  - needs a little fine tuning



 '**************
 ' Flipper Subs
 '**************

 Sub SolLFlipper(Enabled)
     If Enabled Then
         vpPlay "FlipperSu", LeftFlipper:LeftFlipper.RotateToEnd

     Else
         vpPlay "FlipperGiu", LeftFlipper:LeftFlipper.RotateToStart
     End If
 End Sub

 Sub SolRFlipper(Enabled)
     If Enabled Then
         vpPlay "FlipperSu2", RightFlipper:RightFlipper.RotateToEnd
     Else
         vpPlay "FlipperGiu2", RightFlipper:RightFlipper.RotateToStart
     End If
 End Sub


Sub UpdateFlipperLogos
    LFLogo.RotAndTra2 = LeftFlipper.CurrentAngle
    RFlogo.RotAndTra2 = RightFlipper.CurrentAngle
End Sub

Sub SolRampDiverter(enabled)
    if enabled then
        Playsound "DiverterRamp"
        RampDiv.RotateToEnd
        'PrimRampDiv.RotY= 2
        PrimRampDiv.RotY= -3
        PrimRampDiv.TransX = -45
        PrimRampDiv.TransZ= -55
    else
        Playsound "DiverterRamp"
        RampDiv.RotateToStart
        PrimRampDiv.RotY= 12
        PrimRampDiv.TransX = 0
        PrimRampDiv.TransZ= 0
    end if
End Sub

'Sub SolLockRelease(enabled)
'	If enabled then
'		PernoLock.TransY= -29
'		Playsound "DiverterLock"
'		'lockrelease.isdropped = true
'		vlLock.SolExit enabled
'		LockReleaseTimer.Enabled = True
'	End If
'End Sub
'
'Sub LockReleaseTimer_Timer()				'Give LockRelease more time to be down
'	PernoLock.TransY= 0
'	Playsound "DiverterLock"
'	'LockRelease.IsDropped = False
'	LockReleaseTimer.Enabled = False
'End Sub

'Below was a suggested solution for the commented above

Sub SolLockRelease(enabled)
vlLock.SolExit enabled
If enabled then
PernoLock.TransY= -29
Else
PernoLock.TransY= 0
end if
Playsound "DiverterLock"
End Sub


Sub SolKickBack(enabled)
    if enabled then
        CornerKicker.Enabled = True
    else
        KickbackDisableTimer.enabled = true
    End if
End Sub

Sub KickBackDisableTimer_Timer()
    Me.enabled = false
    CornerKicker.enabled = false
End Sub

Sub CornerKicker_Hit()
	ClearBallid
	CornerKicker.DestroyBall
	CreateBallID(CornerKicker)
    CornerKicker.kick 0, 45 'ERA 45
    vpPlay "Rilancio", CornerKicker
End Sub

Sub LeftVelDamp_Hit()
    Activeball.velY = 2
    Activeball.VelX = -2
End Sub

'--------------------------------------------------------
' Give meaningful name to switches and solenoids
'--------------------------------------------------------



' Switches                    |    Status
'-------------------------------------------
'switch 11 unused               'not used
Const swMagGoalieButton  = 12   'installed
Const swStartButton      = 13   'installed
Const swTilt             = 14   'installed
Const swLeftFlipperLane  = 15   'installed
Const swStriker3High     = 16   'installed
Const swRightFlipperLane = 17   'installed
Const swRightOutlane     = 18   'installed
Const swSlamTilt         = 21   'installed
Const swCoinDoor         = 22   'installed
Const swBuyInButton      = 23   'installed
Const swAlwaysClosed     = 24   'not used - always closed
Const swFreeKickTarget   = 25   'installed
Const swKickbackUpper    = 26   'installed
Const swSpinner          = 27   'installed
Const swLightKickback    = 28   'installed
Const swTrough1          = 31   'installed
Const swTrough2          = 32   'installed
Const swTrough3          = 33   'installed
Const swTrough4          = 34   'installed
Const swTrough5          = 35   'installed
Const swTroughStack      = 36   'installed
Const swLightMagGoalie   = 37   'installed
Const swBallShooter      = 38   'installed
Const swGoalTrough       = 41   'installed
Const swGoalPopperOpto   = 42   'installed
Const swGoalIsLeft       = 43   'installed
Const swGoalIsRight      = 44   'installed
Const swTVBallPopper     = 45   'installed
'switch 46 unused               'not used
Const swTravelLaneRolo   = 47   'installed
Const swGoalieTarget     = 48   'installed
Const swSkillShotFront   = 51   'installed
Const swSkillShotCenter  = 52   'installed
Const swSkillShotRear    = 53   'installed
Const swRightEjectHole   = 54   'installed
Const swUpperEjectHole   = 55   'installed
Const swLeftEjectHole    = 56   'installed
Const swRightLaneHi      = 57   'not used
Const swRightLaneLo      = 58   'not used
Const swRollover1        = 61   'installed
Const swRollover2        = 62   'installed
Const swRollover3        = 63   'installed
Const swRollover4        = 64   'installed
Const swTackleSwitch     = 65   'installed  only using 1 long target - I think it might be 3 targets
Const swStriker1         = 66   'installed
Const swStriker2         = 67   'installed
'switch 68 is unused            'not used
Const swLeftRampDiverter = 71   'installed
Const swLeftRampEntrance = 72   'installed
'switch 73 is unused            'not used
Const swLeftRampExit     = 74   'installed
Const swRightRampEntrance= 75   'installed
Const swLockMechLow      = 76   'installed
Const swLockMechHigh     = 77   'installed
Const swRightRampExit    = 78   'installed
Const swLeftJetBumper    = 81   'installed
Const swUpperJetBumper   = 82   'installed
Const swLowerJetBumper   = 83   'installed
Const swLeftSlingshot    = 84   'installed
Const swRightSlingshot   = 85   'installed
Const swKickback         = 86   'installed
Const swUpperLeftLane    = 87   'installed
Const swUpperRightLane   = 88   'installed


Sub TopLeftVelCheck_Hit()
    Activeball.VelY = 1
End Sub

 Sub Trigger1_Hit:ActiveBall.VelZ=0:End Sub
 Sub Trigger2_Hit:ActiveBall.VelZ=0:End Sub

Sub LockRelease_Hit : vpPlay "MetalHit", PernoLock : End Sub



Sub LRHelp_Hit()
    Stopsound "plasticrolling"
    ActiveBall.Velz=0
    Activeball.Velx=0
    ActiveBall.Vely=0
    LRHelp.TimerEnabled=1
End Sub

Sub LRHelp_Timer()
    vpPlay "balldropTopLeft", LRHelp
    LRHelp.TimerEnabled=0
End Sub

Sub RRHelp_Hit()
    Stopsound "plasticrolling"
    RRHelp.TimerEnabled=1
End Sub

Sub RRHelp_Timer()
    vpPlay "balldropBottomRight", RRHelp
    RRHelp.TimerEnabled=0
End Sub

Sub TRHelp_Hit()
    Stopsound "plasticrolling"
    vpPlay "balldropTopLeft", TRHelp
    ActiveBall.Velz=0
    Activeball.Velx=0
    ActiveBall.Vely=0
End Sub

Sub ExitSkill_Hit()
    vpPlay "balldropBottomRight", ExitSkill
End Sub


Sub railsound_Hit()
    vpPlay "metalrolling", ActiveBall
End Sub

Sub railend_Hit()
    StopSound "metalrolling"
    vpPlay "WireRampHit", railend
End Sub

Sub railend1_Hit()
    vpPlay "balldropBottomRight", railend1
End Sub

Sub exit1_Hit()
    StopSound "metalrolling"
    exit1.TimerEnabled=1
End Sub

Sub exit1_Timer()
    vpPlay "balldropTOP", exit1
    exit1.TimerEnabled=0
End Sub
Sub exit2_Hit()
    StopSound "metalrolling"
    exit2.TimerEnabled=1
End Sub

Sub exit2_Timer()
    vpPlay "balldropTOP", exit2
    exit2.TimerEnabled=0
End Sub

Sub exit3_Hit()
    StopSound "metalrolling"
    exit3.TimerEnabled=1
End Sub

Sub exit3_Timer()
    vpPlay "balldropTOP", exit3
    exit3.TimerEnabled=0
End Sub


 '***********
' Update GI
'***********


Dim bulb


Sub UpdateGI(nr,enabled)
 'DOF 200, enabled*-1
 Select Case nr
 Case 0
 For each bulb in GI
 bulb.state=enabled
' GestioneGIWall
 next
 End Select
End Sub



'**************************************
' Fading VPM Lamps VP9 (Reduced/Faster)
'     Based on PD's Fading Lights
' SetLamp 0 is Off
' SetLamp 1 is On
' LampState(x) current state
'**************************************
'**************************************

Dim LampState(200), FadingLevel(200)
Dim FlashSpeedUp(200), FlashSpeedDown(200), FlashMin(200), FlashMax(200), FlashLevel(200)

InitLamps()             ' turn off the lights and flashers and reset them to the default parameters
LampTimer.Interval = 10 'lamp fading speed
LampTimer.Enabled = 1

Sub LampTimer_Timer()
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

Sub InitLamps()
    Dim x
    For x = 0 to 200
        LampState(x) = 0         ' current light state, independent of the fading level. 0 is off and 1 is on
        FadingLevel(x) = 4       ' used to track the fading state
        FlashSpeedUp(x) = 0.5    ' faster speed when turning on the flasher
        FlashSpeedDown(x) = 0.35 ' slower speed when turning off the flasher
        FlashMax(x) = 1          ' the maximum value when on, usually 1
        FlashMin(x) = 0          ' the minimum value when off, usually 0
        FlashLevel(x) = 0        ' the intensity of the flashers, usually from 0 to 1
    Next
End Sub

 Sub UpdateLamps()
    NFadeL 11,  L11
    NFadeL 12,  L12
    NFadeL 13,  L13
    NFadeL 14,  L14
    NFadeL 15,  L15
    NFadeL 16,  L16
    NFadeL 17,  L17
    NFadeL 18,  L18
    NFadeL 21,  L21
    NFadeL 22,  L22
    NFadeL 23,  L23
    NFadeL 24,  L24
    NFadeL 25,  L25
    NFadeL 26,  L26
    NFadeL 27,  L27
    NFadeL 28,  L28
    NFadeL 31,  L31
    NFadeL 32,  L32
    NFadeL 33,  L33
    NFadeL 34,  L34
    NFadeL 35,  L35
    NFadeL 36,  L36
    NFadeL 37,  L37
    NFadeL 38,  L38
    NFadeL 41,  L41
    NFadeL 42,  L42
    NFadeL 43,  L43
    NFadeL 44,  L44
    NFadeLm 45, L45
    NFadeL 45,  l45a
    NFadeLm 46,  L46
    NFadeL 46,  l46a
    NFadeLm 47,  L47
    NFadeL 47,  l47a
    NFadeLm 51, L51
    NFadeL 51,  L51a
    NFadeL 52,  L52
    NFadeL 53,  L53
    NFadeL 54,  L54
    NFadeL 55,  L55
    NFadeL 56,  L56
    NFadeL 57,  L57
    NFadeL 58,  L58
    NFadeL 61,  L61
    NFadeL 62,  L62
    NFadeLm 63, L63
    NFadeL 63,  l63a
    NFadeL 64,  L64
    NFadeLm 65, L65
    NFadeL 65,  L65a
    NFadeL 66,  L66
    NFadeL 67,  L67
    NFadeL 72,  L72
    NFadeL 73,  L73
    NFadeL 74,  L74
    NFadeLm 75,  L75
    NFadeLn 75,  l75a
    NFadeL 75,  L75b
    NFadeL 81,  L81
    NFadeL 82,  L82
    NFadeL 83,  L83
    NFadeL 84,  L84
    'NFadeL 85,  F85

'flashers

    Flash 48, F48
    Flash 68, F68
    Flashm 71, F71A
    Flash 71, F71
    Flash 76, F76
    Flash 77, F77
    Flashm 78, F78A
    Flash 78, F78
    Flash 85, F85
    Flash 86, F86
    FadeModLamp 117, F117,1
    FadeModLamp 119, F119,1
    FadeModLamp 122, F122,1
    FadeModLamp 125, F125,1
    FadeModLamp 126, F126,1
    FadeModLamp 127, F127A,1
    FadeModLamp 127, F127,1
	FadeModLamp 128, F128, 1
	FadeModLamp 128, F128A, 1

	if BrightFlashers then
		'FadeModLamp 126, F126B, 1
		FadeModLamp 117, F117B, 1
		FadeModLamp 119, F119B, 1
		FadeModLamp 128, F128B, 1
		FadeModLamp 128, F128C, 1
		FadeModLamp 127, F127C,1
		FadeModLamp 127, F127B,1
	End if

'    Flashm 128, F128A
'    Flash 128, F128






 End Sub

Sub SetLamp(nr, value)
    If value <> LampState(nr) Then
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

Sub NFadeLm(nr, object) ' used for 2 lights
    Select Case FadingLevel(nr)
        Case 4:object.state = 0
        Case 5:object.state = 1
    End Select
End Sub

Sub NFadeLn(nr, object) ' used for 3 lights
    Select Case FadingLevel(nr)
        Case 4:object.state = 0
        Case 5:object.state = 1
    End Select
End Sub


Sub FadeObj(nr, object, a, b, c, d)
    Select Case FadingLevel(nr)
        Case 2:object.image = d:FadingLevel(nr) = 0 'Off
        Case 3:object.image = c:FadingLevel(nr) = 2 'fading...
        Case 4:object.image = b:FadingLevel(nr) = 3 'fading...
        Case 5:object.image = a:FadingLevel(nr) = 1 'ON
 End Select
End Sub

Sub FadeObjm(nr, object, a, b, c, d)
    Select Case FadingLevel(nr)
        Case 2:object.image = d
        Case 3:object.image = c
        Case 4:object.image = b
        Case 5:object.image = d
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


Sub SetModLamp(nr, value)
    If value > 0 Then
		LampState(nr) = 1
	Else
		LampState(nr) = 0
	End If
	FadingLevel(nr) = value
End Sub

Sub FadeModLamp(nr, object, factor)
	Object.IntensityScale = FadingLevel(nr) * factor/255
	If TypeName(object) = "Light" Then
		Object.State = LampState(nr)
	End If
	If TypeName(object) = "Flasher" Then
		Object.visible = LampState(nr)
	End If
End Sub

'******************
' RealTime Updates
'******************
Set MotorCallback = GetRef("GameTimer")

Sub GameTimer
'   RollingSound
    UpdateFlipperLogos
    UpdateVisuals
End Sub

Sub GatesTimer_Timer()
    'UpdateFlipperLogos
End Sub

Sub UpdateFlipperLogos
    flipperl.RotY = LeftFlipper.CurrentAngle
    flipperr.RotY = RightFlipper.CurrentAngle
    PrimSpinner1.RotZ= -Spinner1.CurrentAngle
End Sub


'****************************************
' B2B Collision by Steely & Pinball Ken
'****************************************
' For use with core.vbs 3.37 or greater to grab BSize variable

Dim tnopb, nosf, iball, cnt, errMessage, B2BOn

'B2BOn = 2 '0=Off, 1=On, 2=AutoDetect
CheckB2B
XYdata.interval = 10 ' <<<<< ADD timer named XYData to table
tnopb = 5   ' <<<<< SET to the "Total Number Of Possible Balls" in play at any one time
nosf = 10   ' <<<<< SET to the "Number Of Sound Files" used / B2B collision volume levels

ReDim CurrentBall(tnopb), BallStatus(tnopb)

For cnt = 0 to ubound(BallStatus) : BallStatus(cnt) = 0 : Next

'****************************************
' B2B AutoDisable for XP x64 Added by Koadic
'****************************************

 Sub CheckB2B           ' Added by Koadic for XP x64 handling
  Dim osver, cpuver, check
  On Error Resume Next
    For x = 0 to 1 : If B2BOn = x Then Exit Sub : End If : Next 'If B2BOn is set manually, then end routine
    Set check = CreateObject("WScript.Shell")
    osver = check.RegRead ("HKLM\Software\Microsoft\Windows NT\CurrentVersion\CurrentVersion")
    cpuver = check.RegRead ("HKLM\SYSTEM\ControlSet001\Control\Session Manager\Environment\Processor_Architecture")
    If osver < 6 and cpuver = "AMD64" Then B2BOn = 0 Else B2BOn = 1 'If OS is XP and 64bit, then disable B2B
    If Err Then B2BOn = 1 'If there is an error in detecting either OS or x32/x64, then default to On
  On Error Goto 0
 End Sub

'======================================================
' <<<<<<<<<<<<<< Ball Identification >>>>>>>>>>>>>>
'======================================================

'******************************
' Destruk's alternative vpmCreateBall for use with B2B Enabled tables
' Core.vbs calls vpmCreateBall when a ball is created from a ball stack
'******************************
 If IsEmpty(Eval("vpmCreateBall"))=false Then Set vpmCreateBall = GetRef("B2BvpmCreateBall")    ' Override the core.vbs and redefine vpmCreateBall

 Function B2BvpmCreateBall(aKicker)
    Dim bsize2:If IsEmpty(Eval("ballsize"))=true Then bsize2 = 25 Else bsize2 = ballsize/2
    For cnt = 1 to ubound(ballStatus)               ' Loop through all possible ball IDs
        If ballStatus(cnt) = 0 Then                 ' If ball ID is available...
            If Not IsEmpty(vpmBallImage) Then       ' Set ball object with the first available ID
                Set CurrentBall(cnt) = aKicker.Createsizedball(bsize2).Image
            Else
                Set CurrentBall(cnt) = aKicker.Createsizedball(bsize2)
            End If
            Set B2BvpmCreateBall = aKicker
            CurrentBall(cnt).uservalue = cnt        ' Assign the ball's uservalue to it's new ID
            ballStatus(cnt) = 1                     ' Mark this ball status active
            ballStatus(0) = ballStatus(0)+1         ' Increment ballStatus(0), the number of active balls
            If B2BOn > 0 Then                       ' If B2BOn is 0, it overrides auto-turn on collision detection
                                                    ' If more than one ball active, start collision detection process
                If ballStatus(0) > 1 and XYdata.enabled = False Then XYdata.enabled = True
            End If
            Exit For                                ' New ball ID assigned, exit loop
        End If
    Next
 End Function

' Use CreateBallID(kickername) to manually create a ball with a BallID
' Can also be used on nonVPM tables (EM or Custom)

 Sub CreateBallID(aKicker)
    Dim bsize2:If IsEmpty(Eval("ballsize"))=true Then bsize2 = 25 Else bsize2 = ballsize/2
    For cnt = 1 to ubound(ballStatus)               ' Loop through all possible ball IDs
        If ballStatus(cnt) = 0 Then                 ' If ball ID is available...
            Set CurrentBall(cnt) = aKicker.Createsizedball(bsize2)      ' Set ball object with the first available ID
            CurrentBall(cnt).uservalue = cnt        ' Assign the ball's uservalue to it's new ID
            ballStatus(cnt) = 1                     ' Mark this ball status active
            ballStatus(0) = ballStatus(0)+1         ' Increment ballStatus(0), the number of active balls
            If B2BOn > 0 Then                       ' If B2BOn is 0, it overrides auto-turn on collision detection
                                                    ' If more than one ball active, start collision detection process
                If ballStatus(0) > 1 and XYdata.enabled = False Then XYdata.enabled = True
            End If
            Exit For                                ' New ball ID assigned, exit loop
        End If
    Next
 End Sub

' Use CreateBallID2(kickername, ballsize) to manually create a custom sized ball with a BallID
' Can also be used on nonVPM tables (EM or Custom)

 Sub CreateBallID2(aKicker, bsize2)                 ' Use to manually create a ball with a BallID with a custom size
    For cnt = 1 to ubound(ballStatus)               ' Loop through all possible ball IDs
        If ballStatus(cnt) = 0 Then                 ' If ball ID is available...
            Set CurrentBall(cnt) = aKicker.Createsizedball(bsize2/2)        ' Set ball object with the first available ID
            CurrentBall(cnt).uservalue = cnt        ' Assign the ball's uservalue to it's new ID
            ballStatus(cnt) = 1                     ' Mark this ball status active
            ballStatus(0) = ballStatus(0)+1         ' Increment ballStatus(0), the number of active balls
            If B2BOn > 0 Then                       ' If B2BOn is 0, it overrides auto-turn on collision detection
                                                    ' If more than one ball active, start collision detection process
                If ballStatus(0) > 1 and XYdata.enabled = False Then XYdata.enabled = True
            End If
            Exit For                                ' New ball ID assigned, exit loop
        End If
    Next
 End Sub

'Call this sub from every kicker that destroys a ball, before the ball is destroyed.

 Sub ClearBallid
    On Error Resume Next                            ' Error handling for debugging purposes
    iball = ActiveBall.uservalue                    ' Get the ball ID to be cleared
    If Err Then Msgbox Err.description & vbCrLf & iball
    ballStatus(iBall) = 0                           ' Clear the ball status
    ballStatus(0) = ballStatus(0)-1                 ' Subtract 1 ball from the # of balls in play
    On Error Goto 0
 End Sub

'=====================================================
' <<<<<<<<<<<<<<<<< XYdata_Timer >>>>>>>>>>>>>>>>>
'=====================================================

'Ball data collection and B2B Collision detection.

ReDim baX(tnopb,4), baY(tnopb,4), baZ(tnopb,4), bVx(tnopb,4), bVy(tnopb,4), TotalVel(tnopb,4)
Dim cForce, bDistance, xyTime, cFactor, id, id2, id3, B1, B2

 Sub XYdata_Timer()
    xyTime = Timer+(XYdata.interval*.001)                       ' xyTime is the system timer plus the current interval time
    If id2 >= 4 Then id2 = 0                                    ' Loop four times and start over
    id2 = id2+1                                                 ' Increment the ball sampler ID
    For id = 1 to ubound(ballStatus)                            ' Loop once for each possible ball
        If ballStatus(id) = 1 Then                              ' If ball is active...
            baX(id,id2) = round(CurrentBall(id).x,2)            ' Sample x-coord
            baY(id,id2) = round(CurrentBall(id).y,2)            ' Sample y-coord
            baZ(id,id2) = round(CurrentBall(id).z,2)            ' Sample z-coord
            bVx(id,id2) = round(CurrentBall(id).velx,2)         ' Sample x-velocity
            bVy(id,id2) = round(CurrentBall(id).vely,2)         ' Sample y-velocity
            TotalVel(id,id2) = (bVx(id,id2)^2 + bVy(id,id2)^2)  ' Calculate total velocity
            If TotalVel(id,id2) > TotalVel(0,0) Then TotalVel(0,0) = int(TotalVel(id,id2))
        End If
    Next
    If ballStatus(0) <= 1 Then XYdata.enabled = False           ' Turn off timer if one ball or less
    If XYdata.interval >= 40 Then B2BOn = 0 : XYdata.enabled = False    ' Auto-shut off
    If Timer > xyTime * 3 Then B2BOn = 0 : XYdata.enabled = False   ' Auto-shut off
    If Timer > xyTime Then XYdata.interval = XYdata.interval+1  ' Increment interval if needed
 End Sub

'=========================================================
'Ball collision (VPX)
'=========================================================

Sub OnBallBallCollision(ball1, ball2, velocity)
    if tablewpc94.VersionMinor > 3 OR tablewpc94.VersionMajor > 10 Then
		PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 2000, Pan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
	else
		PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 2000, Pan(ball1), 0, Pitch(ball1), 0, 0
	end if
End Sub

Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
    Pitch = BallVel(ball) * 20
End Function

Function BallVel(ball) 'Calculates the ball speed
    BallVel = INT(SQR((ball.VelX ^2) + (ball.VelY ^2) ) )
End Function

'=================================================
' <<<<<<<< GetAngle(X, Y, Anglename) >>>>>>>>
'=================================================
Dim Xin,Yin,rAngle,Radit,wAngle
Function Pi:Pi = 4*Atn(1):End Function
Function dSin(degrees)
    dsin = sin(degrees * Pi/180)
    if ABS(dSin) < 0.000001 Then dSin = 0
    if ABS(dSin) > 0.999999 Then dSin = 1 * sgn(dSin)
End Function

 Sub GetAngle(Xin, Yin, wAngle)
    If Sgn(Xin) = 0 Then
        If Sgn(Yin) = 1 Then rAngle = 3 * Pi/2 Else rAngle = Pi/2
        If Sgn(Yin) = 0 Then rAngle = 0
    Else
        rAngle = atn(-Yin/Xin)
    End If
    If sgn(Xin) = -1 Then Radit = Pi Else Radit = 0
    If sgn(Xin) = 1 and sgn(Yin) = 1 Then Radit = 2 * Pi
    wAngle = round((Radit + rAngle),4)
 End Sub

'********************************JimmyFingers Sound Routines**********************************************
Sub arubberposts_Hit(idx)
    dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
    If finalspeed > 14 then
        vpPlay "bump", ActiveBall
    End if
    If finalspeed >= 4 AND finalspeed <= 14 then
        RandomSoundRubber()
    End If
    If finalspeed < 4 AND finalspeed > 1 then
        RandomSoundRubberLowVolume()
    End If
    Dampen 5, .9, 20
End sub

Sub Rubbers_Hit(idx)
    dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
    If finalspeed > 8 then
        vpPlay "bump", ActiveBall
    End if
    If finalspeed >= 1 AND finalspeed <= 7 then
        RandomSoundRubber()
    End If
    Dampen 5, .9, 20
End sub

Sub RandomSoundRubber()
    Select Case Int(Rnd*3)+1
        Case 1 : vpPlay "rubber_hit_1", ActiveBall
        Case 2 : vpPlay "rubber_hit_2", ActiveBall
        Case 3 : vpPlay "rubber_hit_3", ActiveBall
    End Select
End Sub

Sub RandomSoundRubberLowVolume()
    Select Case Int(Rnd*3)+1
        Case 1 : vpPlay "rubber_hit_1_low", ActiveBall
        Case 2 : vpPlay "rubber_hit_2_low", ActiveBall
        Case 3 : vpPlay "rubber_hit_3_low", ActiveBall
    End Select
End Sub

Sub Dampen(dt,df,r)                     'dt is threshold speed, df is dampen factor 0 to 1 (higher more dampening), r is randomness
    Dim dfRandomness
    r=cint(r)
    dfRandomness=INT(RND*(2*r+1))
    df=df+(r-dfRandomness)*.01
    If ABS(activeball.velx) > dt Then activeball.velx=activeball.velx*(1-df*(ABS(activeball.velx)/100))
    If ABS(activeball.vely) > dt Then activeball.vely=activeball.vely*(1-df*(ABS(activeball.vely)/100))
End Sub

Sub LeftFlipper_Collide(parm)
 	RandomSoundFlipper()
End Sub

Sub RightFlipper_Collide(parm)
 	RandomSoundFlipper()
End Sub

Sub RandomSoundFlipper()
	Select Case Int(Rnd*3)+1
		Case 1 : PlaySound "flip_hit_1"
		Case 2 : PlaySound "flip_hit_2"
		Case 3 : PlaySound "flip_hit_3"
	End Select
End Sub


'*****************************************
'           FLIPPER SHADOWS
'*****************************************

sub FlipperTimer_Timer()
    FlipperLSh.RotZ = LeftFlipper.currentangle
    FlipperRSh.RotZ = RightFlipper.currentangle

End Sub

'*****************************************
'           BALL SHADOW
'*****************************************
Dim BallShadow
BallShadow = Array (BallShadow1,BallShadow2,BallShadow3,BallShadow4,BallShadow5)
Dim tnob:tnob = 5
'
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
        If BOT(b).X < tablewpc94.Width/2 Then
            BallShadow(b).X = ((BOT(b).X) - (Ballsize/6) + ((BOT(b).X - (tablewpc94.Width/2))/7)) + 10
        Else
            BallShadow(b).X = ((BOT(b).X) + (Ballsize/6) + ((BOT(b).X - (tablewpc94.Width/2))/7)) - 10
        End If
        ballShadow(b).Y = BOT(b).Y + 20
        If BOT(b).Z > 20 Then
            BallShadow(b).visible = 1
        Else
            BallShadow(b).visible = 0
        End If
    Next
End Sub


'******************************************
' Use RollingSoundTimer to call div subs
'******************************************


Sub RollingSoundTimer_Timer()
    RollingSound
End Sub


'****************************************
' JimmyFingers Enhanced Ball Rolling Script (Extension of Rascal's Original)
'****************************************

ReDim BRVeloY(tnopb), BRVeloX(tnopb), rolling(tnopb), rollingfast(tnopb)
Dim b
b = 0

RollingSoundTimer.Interval = 60/tnopb

Sub RollingSound()
    B = B + 1
    If B > tnopb Then B = 1
    If BallStatus(b) = 0 Then Exit Sub

    BRVeloY(b) = Cint(CurrentBall(b).VelY)
    BRVeloX(b) = Cint(CurrentBall(b).VelX)
    If((ABS(BRVeloY(b))> 3 AND (ABS(BRVeloY(b))< 10) or (ABS(BRVeloX(b) )> 3 AND (ABS(BRVeloX(b))< 10)))) Then
        If rolling(b) = True then
            Exit Sub
        Else
            If rollingfast(b) = True then
                StopSound "JF_rollingfaster"
                rollingfast(b) = False
            Else
                rolling(b) = True
                Select Case Int(Rnd*3)+1
                    Case 1 : vpPlay "JF_roll1", CurrentBall(b)
                    Case 2 : vpPlay "JF_roll2", CurrentBall(b)
                    Case 3 : vpPlay "JF_roll3", CurrentBall(b)
                End Select
            End If
        End If
    ElseIf (ABS(BRVeloY(b) )> 10 or ABS(BRVeloX(b) )> 10) Then
        If rollingfast(b) = True then
            Exit Sub
        Else
            If rolling(b) = True then
                StopSound "JF_roll1"
                StopSound "JF_roll2"
                StopSound "JF_roll3"
                rolling(b) = False
            Else
                rollingfast(b) = True
                vpPlay "JF_rollingfaster", CurrentBall(b)
            End If
        End If
    Else
        If rolling(b) = True Then
            StopSound "JF_roll1"
            StopSound "JF_roll2"
            StopSound "JF_roll3"
            rolling(b) = False
        ElseIf rollingfast(b) = True Then
            StopSound "JF_rollingfaster"
            rollingfast(b) = False
        End If
    End If
End Sub

Sub StopRollingSound()
    StopSound "JF_roll1"
    StopSound "JF_roll2"
    StopSound "JF_roll3"
    StopSound "JF_rollingfaster"
End Sub

'REGISTRY LOCATIONS ***************************************************************************************************************************************

 Const optOpenAtStart   = &H000001
 Const optDMDRotation   = &H000002
 Const optDMDHidden     = &H000004
 Const optBallGI        = &H000008
 Const optController    = &H000010
 Const optB2BEnable     = &H000100
 Const optGoalieSpeed   = &H001000
 Const optBallImage     = &H010000
 Const optFBSounds      = &H100000

'OPTIONS MENU *********************************************************************************************************************************************

 Dim TableOptions, TableName, optReset
 Private vpmShowDips1, vpmDips1

 Sub InitializeOptions
    TableName="WCS94"                       'Replace with your descriptive table name, it will be used to save settings in VPReg.stg file
    Set vpmShowDips1 = vpmShowDips                              'Reassigns vpmShowDips to vpmShowDips1 to allow usage of default dips menu
    Set vpmShowDips = GetRef("TableShowDips")                   'Assigns new sub to vmpShowDips
    TableOptions = LoadValue(TableName,"Options")               'Load saved table options
    Set Controller = CreateObject("VPinMAME.Controller")        'Load vpm controller temporarily so options menu can be loaded if needed
    If TableOptions = "" Or optReset Then                       'If no existing options, reset to default through optReset, then open Options menu
        TableOptions = DefaultOptions                           'clear any existing settings and set table options to default options
        TableShowOptions
    ElseIf (TableOptions And optOpenAtStart) Then               'If Enable Next Start was selected then
        TableOptions = TableOptions - optOpenAtStart            'clear setting to avoid future executions
        TableShowOptions
    Else
        TableSetOptions
    End If
    Set Controller = Nothing                                    'Unload vpm controller so selected controller can be loaded
 End Sub

 Private Sub TableShowDips
    vpmShowDips1                                                'Show original Dips menu
    TableShowOptions                                            'Show new options menu
'   TableShowOptions2                                           'Add more options menus...
 End Sub

 Private Sub TableShowOptions                   'New options menu, additional menus can be added as well, just follow similar format and add call to TableShowDips
   Dim oldOptions : oldOptions = TableOptions
    If Not IsObject(vpmDips1) Then              'If creating an additional menus, need to declare additional vpmDips variables above (ex. vpmDips2 and TableOptions2, etc.)
        Set vpmDips1 = New cvpmDips
        With vpmDips1
            .AddForm 530, 250, "TABLE OPTIONS MENU"
            .AddFrameExtra 0,0,105,"Controller Selection*",3*optController, Array("Visual PinMame", 1*optController, "UVP", 2*optController,_
                "B2S Server", 3*optController)
            .AddFrameExtra 0,60,105,"DMD Options*",0, Array("Rotate DMD", optDMDRotation, "Hide DMD", optDMDHidden)
            .AddFrameExtra 0,106,105,"B2B Options",3*optB2BEnable, Array("Force Disable", 0*optB2BEnable, "Force Enable", 1*optB2BEnable, "Auto Detect", 2*optB2BEnable)
            .AddLabel 5,166,100,15,"* Requires restart"

            .AddFrameExtra 125,0,105,"Goalie Speed*",3*optGoalieSpeed, Array("Fast", 1*optGoalieSpeed, "Normal", 2*optGoalieSpeed,_
                "Slow", 3*optGoalieSpeed)
            .AddFrameExtra 125,60,105,"Ball Image",3*optBallImage, Array("Black/White", 0*optBallImage, "Teal/White", 1*optBallImage,_
                "Design", 2*optBallImage)
            .AddChkExtra 130,123,150, Array("Disable Ball Shading (FPS+)", optBallGI)
            .AddChkExtra 130,138,150, Array("Disable Mech Sounds", optFBSounds)
            .AddChkExtra 130,166,105, Array("Enable Next Start", optOpenAtStart)
        End With
    End If
    TableOptions = vpmDips1.ViewDipsExtra(TableOptions)
    SaveValue TableName,"Options",TableOptions
    TableSetOptions
 End Sub

 Dim BallImage, BallType, BallGI, GoalieSpeed
 BallImage = Array("soccerball", "soccerball2", "soccerball3")

 Sub TableSetOptions        'define required settings before table is run
    ROL = (TableOptions And optDMDRotation)\optDMDRotation
    HIDDEN = (TableOptions And optDMDHidden)\optDMDHidden
    cController = ((TableOptions And (3*optController))\optController)
    B2Bon = ((TableOptions And (3*optB2BEnable))\optB2BEnable)
    BallType = ((TableOptions And (3*optBallImage))\optBallImage)
    BallGI = (TableOptions And optBallGI)\optBallGI
    GoalieSpeed = ((TableOptions And (3*optGoalieSpeed))\optGoalieSpeed)
    Select Case GoalieSpeed
        Case 1:GoalieSpeed = 60
        Case 2:GoalieSpeed = 80
        Case 3:GoalieSpeed = 120
    End Select
    If IsObject(mGoalie) Then mGoalie.length = GoalieSpeed
    If cbool(TableOptions AND optFBSounds) Then
        FFBSounds = FeedbackSounds
        BallRel = ""
        SolOn = ""
        Popper = ""
    Else
        FFBSounds = Empty
        BallRel = "BallRel"
        SolOn = "SolOn"
        Popper = "Popper"
    End If
 End Sub

 Sub UpdateVisuals
    If SoccerBall.image <> BallImage(BallType) Then SoccerBall.image = BallImage(BallType)
    If SoccerBallLight.TopVisible = cbool(BallGI) Then SoccerBallLight.TopVisible = Not cbool(BallGI)
 End Sub

 Dim FFBSounds
 Sub vpPlay(sound, tableobj)
  Dim x
    If Not IsEmpty(FFBSounds) Then  'If FFBSounds is assigned to the feedbacksounds array... aka FeedBack sounds turned OFF
        For x = 0 to Ubound(FFBSounds)                  'Loop through all sounds in the array
            If lcase(FFBSounds(x)) = lcase(sound) Then  'Check to see if sound is present in array, and if so
                Exit Sub                                'Exit the sub as no sound should then be played
            End If
        Next
    End If 'If sound isn't found, then play sound as normal...
    Debug.print sound & Pan(tableobj) & AudioFade(tableobj)

    If tablewpc94.VersionMinor > 3 OR VersionMajor > 10 Then
        PlaySound sound, 1, 1, Pan(tableobj), 0,0,0, 1, AudioFade(tableobj)
    Else
        PlaySound sound, 1, 1, Pan(tableobj)
    End If
    'VPX 10.4 only

 End Sub

function AudioFade(ball)
    Dim tmp
    tmp = ball.y * 2 / tablewpc94.height-1
    If tmp > 0 Then
        AudioFade = Csng(tmp ^10)
    Else
        AudioFade = Csng(-((- tmp) ^10) )
    End If
End Function

Function Pan(tobj) ' Calculates the pan for a tableobj based on the X position on the table. "table1" is the name of the table
    Dim tmp
    tmp = tobj.x * 2 / tablewpc94.width-1
    If tmp > 0 Then
        Pan = Csng(tmp ^10)
    Else
        Pan = Csng(-((- tmp) ^10) )
    End If
End Function

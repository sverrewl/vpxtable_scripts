'   _____  _             ___  ___              _      _
'  |_   _|| |            |  \/  |             | |    (_)
'    | |  | |__    ___   | .  . |  __ _   ___ | |__   _  _ __    ___
'    | |  | '_ \  / _ \  | |\/| | / _` | / __|| '_ \ | || '_ \  / _ \ 
'    | |  | | | ||  __/  | |  | || (_| || (__ | | | || || | | ||  __/
'    \_/  |_| |_| \___|  \_|  |_/ \__,_| \___||_| |_||_||_| |_| \___|
'
'
'______        _      _                  __   ______  _         _             _
'| ___ \      (_)    | |                / _|  | ___ \(_)       | |           | |
'| |_/ / _ __  _   __| |  ___     ___  | |_   | |_/ / _  _ __  | |__    ___  | |_
'| ___ \| '__|| | / _` | / _ \   / _ \ |  _|  |  __/ | || '_ \ | '_ \  / _ \ | __|
'| |_/ /| |   | || (_| ||  __/  | (_) || |    | |    | || | | || |_) || (_) || |_
'\____/ |_|   |_| \__,_| \___|   \___/ |_|    \_|    |_||_| |_||_.__/  \___/  \__|
'
'
'  _    _  _  _  _  _                           __   _____  _____  __
' | |  | |(_)| || |(_)                         /  | |  _  ||  _  |/  |
' | |  | | _ | || | _   __ _  _ __ ___   ___   `| | | |_| || |_| |`| |
' | |/\| || || || || | / _` || '_ ` _ \ / __|   | | \____ |\____ | | |
' \  /\  /| || || || || (_| || | | | | |\__ \  _| |_.___/ /.___/ /_| |_
'  \/  \/ |_||_||_||_| \__,_||_| |_| |_||___/  \___/\____/ \____/ \___/


'The Machine (Bride of Pinbot) Williams 1991 Rev1.3 for VP10 (Requires VP10.3 final release or greater to play)

'***The very talented BOP development team.***
'Original VP10 beta by "Unclewilly and completed by wrd1972"
'Additional scripting assistance by "cyberpez", "Rothbauerw".
'Clear ramps, wire ramp primitives, pop-bumper caps and flasher domes, shadows and flipper prims by "Flupper"
'Additional 3D work by "Dark" and "Cyberpez"
'Space shuttle toy and sidewall graphics by "Cyberpez"
'Face rotation scripting by "KieferSkunk/Dorsola"
'High poly helmet and re-texture by "Dark"
'Original bride helmet, plastics and playfield redraw by "Jc144"
'Desktop view, scoring dials aand background by "32Assassin"
'Full funtion ball trough by "cyberpez"
'Lighting by wrd1972
'Flashers by "wrd1972"
'HMLF table physics by "wrd1972"
'Flipper physics by "Rothbauerw"
'Modulated GI lighting by "nFozzy"
'LED color GI lighting by "cyberpez"
'DOF and controller by "Arngrim"
'***Many...many thanks to you all for the tremenedous efforts and hard work for making this awesome table a reality. I just cant thank you all enough.***

' Thalamus 2018-07-24
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Changed UseSolenoids=1 to 2
' No special SSF tweaks yet.

Option Explicit
Randomize
Dim OptReset



'*************************************************************************************************************************
'*************************************************************************************************************************
'______                            _     _ _               _   _            _____     _     _
'| ___ \                          | |   | (_)             | | | |          |_   _|   | |   | |
'| |_/ /___ ______ ___ _ __   __ _| |__ | |_ _ __   __ _  | |_| |__   ___    | | __ _| |__ | | ___
'|    // _ \______/ _ \ '_ \ / _` | '_ \| | | '_ \ / _` | | __| '_ \ / _ \   | |/ _` | '_ \| |/ _ \ 
'| |\ \  __/     |  __/ | | | (_| | |_) | | | | | | (_| | | |_| | | |  __/   | | (_| | |_) | |  __/
'\_| \_\___|      \___|_| |_|\__,_|_.__/|_|_|_| |_|\__, |  \__|_| |_|\___|   \_/\__,_|_.__/|_|\___|
'                                                   __/ |
'                                                  |___/
' _____       _   _                   _____ _             _                ___  ___
'|  _  |     | | (_)                 /  ___| |           | |               |  \/  |
'| | | |_ __ | |_ _  ___  _ __  ___  \ `--.| |_ __ _ _ __| |_ _   _ _ __   | .  . | ___ _ __  _   _
'| | | | '_ \| __| |/ _ \| '_ \/ __|  `--. \ __/ _` | '__| __| | | | '_ \  | |\/| |/ _ \ '_ \| | | |
'\ \_/ / |_) | |_| | (_) | | | \__ \ /\__/ / || (_| | |  | |_| |_| | |_) | | |  | |  __/ | | | |_| |
' \___/| .__/ \__|_|\___/|_| |_|___/ \____/ \__\__,_|_|   \__|\__,_| .__/  \_|  |_/\___|_| |_|\__,_|
'      | |                                                         | |
'      |_|
'                                                        |_|
'Step 1. Uncomment (remove apostrophe at beginning of line 76) to re-enable the Table Options Startup Menu.
'Step 2  Re-start the table (menu should now appear).
'Step 3. Re-comment (replace apostrophe at beginning of line 76) after table is restarted.

'OptReset = 1
'*************************************************************************************************************************
'*************************************************************************************************************************


'RomName
cGameName = "bop_l7"

Const BallSize = 25  'Ball radius
Const ballmass = 1.7   'Ball mass
const UseVPMModSol = True

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Dim cGameName, ShipMod, cheaterpost, sidewalls, helmetreflections, Musicsnippet, Prevgameover
MusicSnippet = 0

Dim DesktopMode: DesktopMode = Table.ShowDT

If DesktopMode = True Then 'Show Desktop components
    frontlockbar.visible=1
    rearbar.visible=1
    pSidewall.Z = -55

Else
    frontlockbar.visible=0
    rearbar.visible=0
    pSidewall.Z = -55
End if





LoadVPM "01560000", "WPC.VBS", 3.26
'Variables
Dim bsSaucer, bsLEye, bsREye, bsMouth, bsSS, mFace, xx, bump1, bump2, bump3
Dim mechHead, headAngle, prevHeadAngle, currentFace
Dim MaxBalls

MaxBalls=3
Const UseSolenoids = 2
Const UseLamps = 0
Const UseSync = 1
Const HandleMech = 0

'Standard Sounds
Const SSolenoidOn = "Solenoid"
Const SSolenoidOff = ""
Const SCoin = "CoinIn"

' Setup the lightning according to the nightday slider
DayNightStuff
Sub DayNightStuff()
    Dim ii
    On Error Resume Next
    for each ii in GI_rear: ii.intensity = ii.intensity + (100 - Table.nightday)/10: next
    for each ii in GI_front: ii.intensity = ii.intensity + (100 - Table.nightday)/10: next
    'for each ii in aFlashers:ii.opacity=ii.opacity + (100 - Table.nightday)^2:next
    for each ii in aLampsAll:ii.opacity=ii.opacity + (100 - Table.nightday)/10:next
    for each ii in aLampsAll:ii.Intensity=ii.Intensity + (100 - Table.nightday)/10:next
End Sub

'Table Init
Sub Table_Init
    vpmInit Me
    With Controller
        .GameName =  cGameName
        If Err Then MsgBox "Can't start Game " & cGameName & vbNewLine & Err.Description:Exit Sub
        .SplashInfoLine = "The Machine BOP, Williams 1991" & vbNewLine & "by wrd1972"
        .HandleKeyboard = 0
        .ShowTitle = 0
        .ShowDMDOnly = 1
        .ShowFrame = 0
        .HandleMechanics = 0
        .Hidden = 1
        On Error Resume Next
        .Run GetPlayerHWnd
        If Err Then MsgBox Err.Description
        On Error Goto 0
    End With

'	ChangeBats(ChooseBats)

    'Nudging
    vpmNudge.TiltSwitch=14
    vpmNudge.Sensitivity=1
    vpmNudge.TiltObj=Array(Bumper1b,Bumper2b,Bumper3b,LeftSlingshot,RightSlingshot,UpperSlingShot)

      Set bsLEye = New cvpmBallStack
      With bsLEye
         .InitSaucer sw63,63, 180, 5
        .InitExitSnd SoundFX("fx_popper",DOFContactors), SoundFX("Solenoid",DOFContactors)
       .InitAddSnd "fx_kicker_catch"
      End With

      Set bsREye = New cvpmBallStack
      With bsREye
         .InitSaucer sw64,64, 180, 5
         .InitExitSnd SoundFX("fx_popper",DOFContactors), SoundFX("Solenoid",DOFContactors)
        .InitAddSnd "fx_kicker_catch"
      End With

      Set bsMouth = New cvpmBallStack
      With bsMouth
         .InitSaucer sw65,65, 180, 5
         .InitExitSnd SoundFX("fx_popper",DOFContactors), SoundFX("Solenoid",DOFContactors)
        .InitAddSnd "fx_kicker_catch"
      End With

      Set bsSS = New cvpmBallStack
      With bsSS
         .InitSaucer SSLaunch,31, 0, 40
         .KickForceVar = 2
         .InitExitSnd SoundFX("Solenoid",DOFContactors), SoundFX("Solenoid",DOFContactors)
      End With

      Set mechHead = New cvpmMech
      With mechHead
         .MType = vpmMechOneDirSol + vpmMechCircle + vpmMechLinear + vpmMechSlow
         .Sol1 = 28
         .Sol2 = 27
         .Length = 60 * 10   ' 10 seconds to cycle through all faces.
         .Steps = 360 * 4    ' 4 wheel rotations for one full head rotation
         .Callback = GetRef("HeadMechCallback")
         .Start
      End With

    '**Main Timer init
    PinMAMETimer.Enabled = 1

     PrevGameOver = 0

    SetOptions

    CheckMaxBalls 'Allow balls to be created at table start up


End Sub

Sub Table_Paused:Controller.Pause = 1:End Sub
Sub Table_unPaused:Controller.Pause = 0:End Sub
Sub Table_Exit:Controller.Stop:End Sub


'*****Keys
Sub Table_KeyDown(ByVal keycode)
     If keycode = plungerkey then plunger.PullBack:PlaySound "plungerpull"
     'If Keycode = RightFlipperKey then SolRFlipper 1
     'If Keycode = LeftFlipperKey then SolLFlipper 1
     If keycode = LeftTiltKey Then LeftNudge 80, 1, 20
     If keycode = RightTiltKey Then RightNudge 280, 1, 20
     If keycode = CenterTiltKey Then CenterNudge 0, 1, 25

    '************************   Start Ball Control 1/3
        if keycode = 46 then                ' C Key
            If contball = 1 Then
                contball = 0
            Else
                contball = 1
            End If
        End If
        if keycode = 48 then                'B Key
            If bcboost = 1 Then
                bcboost = bcboostmulti
            Else
                bcboost = 1
            End If
        End If
        if keycode = 203 then bcleft = 1        ' Left Arrow
        if keycode = 200 then bcup = 1          ' Up Arrow
        if keycode = 208 then bcdown = 1        ' Down Arrow
        if keycode = 205 then bcright = 1       ' Right Arrow
    '************************   End Ball Control 1/3
    If vpmKeyDown(keycode) Then Exit Sub
End Sub


Sub Table_KeyUp(ByVal keycode)

     If keycode = plungerkey then plunger.Fire:PlaySound "Plunger2"
     'If Keycode = RightFlipperKey then SolRFlipper 0
     'If Keycode = LeftFlipperKey then SolLFlipper 0

    '************************   Start Ball Control 2/3
    if keycode = 203 then bcleft = 0        ' Left Arrow
    if keycode = 200 then bcup = 0          ' Up Arrow
    if keycode = 208 then bcdown = 0        ' Down Arrow
    if keycode = 205 then bcright = 0       ' Right Arrow
    '************************   End Ball Control 2/3
     If vpmKeyUp(keycode) Then Exit Sub
End sub



'************************   Start Ball Control 3/3
Sub StartControl_Hit()
    Set ControlBall = ActiveBall
    contballinplay = true
End Sub

Sub StopControl_Hit()
    contballinplay = false
End Sub

Dim bcup, bcdown, bcleft, bcright, contball, contballinplay, ControlBall, bcboost
Dim bcvel, bcyveloffset, bcboostmulti

bcboost = 1     'Do Not Change - default setting
bcvel = 4       'Controls the speed of the ball movement
bcyveloffset = -0.01    'Offsets the force of gravity to keep the ball from drifting vertically on the table, should be negative
bcboostmulti = 3    'Boost multiplier to ball veloctiy (toggled with the B key)

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
'************************   End Ball Control 3/3

 'Solenoids
      SolCallback(1) = "kisort"
       SolCallback(2) = "KickBallToLane"
       SolCallback(3) = "SolKickout"
       SolCallback(4) = "vpmSolGate Gate3,0,"
       SolCallback(5) = "Solss"
       SolCallback(6) = "solBallLockPost"
       SolCallback(7) = "vpmSolSound SoundFX(""fx_Knocker"",DOFKnocker),"
       SolCallback(8) = "bsMouth.SolOut"
       SolCallback(15) = "bsleye.SolOut"
       SolCallback(16) = "bsREye.SolOut"
       REM SolCallback(17) = "Flash17" 'Billions
      REM SolCallback(18) = "SetFlash 118," 'Left ramp
      REM SolCallback(19) = "Flash19" 'jackpot
      REM SolCallback(20) = "Flash20" 'SkillShot
      REM SolCallback(21) = "SetDome1" 'Left Helmet
      REM SolCallback(22) = "SetDome2" 'Right Helmet
      REM SolCallback(23) = "SetDome3" 'Jets Enter
      REM SolCallback(24) = "SetDome4" 'Left Loop

Sub SolKickout(enabled)
    If Enabled then
        sw46k.kick 22.5,45
        sw46.enabled= 0
        vpmtimer.addtimer 600, "sw46.enabled= 1'"
        Playsound SoundFx("fx_Popper",DOFContactors)
    End If
End Sub

Sub sw46_Hit()
    Playsound "fx_vuk_enter"
End Sub

Sub sw46k_Hit()
    Playsound "fx_kicker_catch"
    Controller.Switch(46) = 1
End Sub

Sub sw46k_UnHit()
    Controller.Switch(46) = 0
End Sub


Sub SolMotor(Enabled)
    Light1.State = enabled
End Sub

Sub SolRelay(Enabled)
    Light2.State = enabled
End Sub


  '**************
 ' Solenoid Subs
 '**************

   Sub SolSS(enabled)
    If Enabled then
        bsSS.ExitSol_On:SSKick.transZ = 20:EMPos = 20:SSLaunch.TimerEnabled = 1
    end if
   End Sub

Dim EMPos
    EMPos = 0
  Sub SSLaunch_timer()
    EmPos = EmPos - 1
    SSKick.transZ = EmPos
    If EmPos = 0 then SSLaunch.TimerEnabled = 0
  End Sub

   Sub solBallLockPost(enabled)
    If Enabled then BL.IsDropped=1:BLP.transY = -90:BL.TimerEnabled = 1
   End Sub

    Sub BL_Timer()
        BL.IsDropped=0:BLP.transY = 0:BL.TimerEnabled = 0
    End Sub

'Flashers
REM Sub Flash17(enabled)
    REM Setlamp 117, enabled
REM End Sub

REM Sub Flash19(enabled)
    REM Setlamp 119, enabled
REM End Sub

REM Sub Flash20(enabled)
    REM Setlamp 120, enabled
REM End Sub


'***********************************************
'**************
'ConstantUpdates
'**************

Sub GameTimer_Timer()
    UpdateGatesSpinners
End Sub

Dim Pi, GateSpeed
Pi = Round(4*Atn(1),6)
GateSpeed = 0.5

Dim smGateOpen,smGateAngle:smGateOpen=0:smGateAngle=0
Sub smlGate_Hit():smGateOpen=1:smGateAngle=0:End Sub

Dim Gate2Open,Gate2Angle:Gate2Open=0:Gate2Angle=0
Sub Gate2_Hit():Gate2Open=1:Gate2Angle=0:End Sub

Dim Gate4Open,Gate4Angle:Gate4Open=0:Gate4Angle=0
Sub Gate4_Hit():Gate4Open=1:Gate4Angle=0:End Sub

Dim Gate5Open,Gate5Angle:Gate5Open=0:Gate5Angle=0
Sub Gate5_Hit():Gate5Open=1:Gate5Angle=0:End Sub

Sub UpdateGatesSpinners
    SpinnerT1.RotX = -(sw51.currentangle)
    pSpinnerRod.TransX = sin( (sw51.CurrentAngle+180) * (2*PI/360)) * 5
    pSpinnerRod.TransY = sin( (sw51.CurrentAngle- 90) * (2*PI/360))
    If Gate3.currentangle > 70 Then
        Gate3P.RotZ = -90
    Else
        Gate3P.RotZ = -(Gate3.currentangle+20)
    End If
    Spinner2P.Rotx = -(spinner2.currentangle-90)
    Spinner1P.RotX = -(spinner1.currentangle-90)

    If Gate2Open Then
        If Gate2Angle < Gate2.currentangle Then:Gate2Angle=Gate2.currentangle:End If
        If Gate2Angle > 5 and Gate2.currentangle < 5 Then:Gate2Open=0:End If
        If Gate2Angle > 70 Then
            Gate2P.RotZ = -90
        Else
            Gate2P.RotZ = -(Gate2Angle+20)
            Gate2Angle=Gate2Angle - GateSpeed
        End If
    Else
        if Gate2Angle > 0 Then
            Gate2Angle = Gate2Angle - GateSpeed
        Else
            Gate2Angle = 0
        End If
        Gate2P.RotZ = -(Gate2Angle + 20)
    End If

    If Gate4Open Then
        If Gate4Angle < Gate4.currentangle Then:Gate4Angle=Gate4.currentangle:End If
        If Gate4Angle > 5 and Gate4.currentangle < 5 Then:Gate4Open=0:End If

        Gate4P.RotX = 90 - Gate4Angle
        Gate4Angle=Gate4Angle - GateSpeed

    Else
        if Gate4Angle > 0 Then
            Gate4Angle = Gate4Angle - GateSpeed
        Else
            Gate4Angle = 0
        End If
        Gate4P.RotX = 90 - Gate4Angle
    End If

    If Gate5Open Then
        If Gate5Angle < Gate5.currentangle Then:Gate5Angle=Gate5.currentangle:End If
        If Gate5Angle > 5 and Gate5.currentangle < 5 Then:Gate5Open=0:End If
        If Gate5Angle > 70 Then
            Gate5P.RotZ = 90
        Else
            Gate5P.RotZ = (Gate5Angle+20)
            Gate5Angle=Gate5Angle - GateSpeed
        End If
    Else
        if Gate5Angle > 0 Then
            Gate5Angle = Gate5Angle - GateSpeed
        Else
            Gate5Angle = 0
        End If
        Gate5P.RotZ = Gate5Angle + 20
    End If

    If smGateOpen Then
        If smGateAngle < smlGate.currentangle Then:smGateAngle=smlGate.currentangle:End If
        If smGateAngle > 5 and smlGate.currentangle < 5 Then:smGateOpen=0:End If
        If smGateAngle > 70 Then
            smlGateP.RotZ = 90
        Else
            smlGateP.RotZ = (smGateAngle+20)
            smGateAngle=smGateAngle - GateSpeed
        End If
    Else
        if smGateAngle > 0 Then
            smGateAngle = smGateAngle - GateSpeed
        Else
            smGateAngle = 0
        End If
        smlGateP.RotZ = smGateAngle + 20
    End If

    Gate6P.RotX = (Gate6.currentangle+90)
    'BatLeft.ObjRotz = LeftFlipper.Currentangle
    'BatRight.ObjRotz = RightFlipper.Currentangle
    pFaceDiverter.RotX = div.currentangle / 2
End Sub


'***********************************************
'**************
' Flipper Subs
'**************

 SolCallback(sLRFlipper) = "SolRFlipper"
 SolCallback(sLLFlipper) = "SolLFlipper"

 '******************************************
'Added by JF
'******************************************

'******************************************
' Use FlipperTimers to call div subs
'******************************************

Sub SolLFlipper(Enabled)
     If Enabled Then
         PlaySound SoundFx("lflip",DOFFlippers)
         LeftFlipper.RotateToEnd
     Else
         PlaySound SoundFx("lflipd",DOFFlippers)
         LeftFlipper.RotateToStart
     End If
 End Sub

Sub SolRFlipper(Enabled)
    If enabled then
         PlaySound SoundFx("rflip",DOFFlippers)
         RightFlipper.RotateToEnd
     Else
         PlaySound SoundFx("rflipd",DOFFlippers)
         RightFlipper.RotateToStart
     End If
 End Sub

 '-----------------
' Head/Face rotation script
' cvpmMech-based script by KieferSkunk/Dorsola, May 2015
'-----------------
'
' MOTOR ASSEMBLY (notes from BoP manual)
'
' Motor has a single peg on it at the outside of a disk, which intersects
' with a plus-shaped guide on the head box.  The peg pushes into a slot on
' the guide and rotates the box 90 degrees - during that time, the disk has
' also rotated 90 degrees.  The remaining 270 degrees, the wheel just free-
' spins.  This type of motion causes the face to rotate slowest when the pin
' is entering or exiting the slot (0/90 degrees), and fastest when the peg is
' halfway through the rotation (45 degrees).  A cosine function will simulate
' this speed curve quite nicely.
'
' The mech handler will do 4 full 360-degree rotations to simulate the drive
' wheel through an entire head-box rotation.  The callback will then determine
' the head-box's current position by simulating the peg.
' ---
'
' FACE SWITCH ASSEMBLY:(notes from BoP manual)
'
' Switch 67 is the head position switch.  The switch should be closed when in
' an indentation on the head box bottom plate, and open at all other times.
' Each indent is a concave surface to allow the switch to travel smoothly, so
' we can assume that the switch is closed when at least halfway into the maximum
' depth of an indent.  There are indents on only three sides (same sides as
' Faces 1, 2 and 4) - the last side (Face 3) doesn't have an indent, so the
' switch stays open on this side.  When looking at the head top-down (on playfield),
' this switch is on the right side, which means Face 4 is facing upward when the
' switch is stuck open.  The size of the indents and an assumption about their
' concavity leads me to believe the switch should be closed for about 10 degrees
' either side of center.

headAngle = 0
prevHeadAngle = 0
currentFace = 1

Function headIsNear(target)
    headIsNear = (headAngle >= target - 10) AND (headAngle <= target + 10)
End Function

Sub HeadMechCallback(aNewPos, aSpeed, aLastPos)
    headAngle = Fix(aNewPos / 360) * 90        ' Get integer position for current face.
   Dim wheelPos : wheelPos = aNewPos - (headAngle * 4)   ' What position is the wheel in?
   ' Wheel position > 270 = head is 90 degrees further along than original base calc.
   If (wheelPos >= 270) Then
      headAngle = headAngle + 90
    ElseIf (wheelPos >= 180) Then
      ' Calculate how far along into rotation the head is.
     ' Since our goal is to get slow movement at both ends of the arc, we need a
     ' Cosine over 180 degrees (Pi radians), so this formula takes care of that.
     ' Also, Cosine varies between -1 and 1, so add 1 to the value (range 0-2) and
     ' cut it in half to get correct partial angle with correct accel curve.
     Dim wheelRads : wheelRads = ((wheelPos - 180) * 2 + 180) * 0.0174532925
      headAngle = headAngle + 45 * (1 + Cos(wheelRads))
    End If

    ' No more to do if the head hasn't changed position.
   If headAngle = prevHeadAngle Then
        Exit Sub
    End If

    ' Head has moved.
   prevHeadAngle = headAngle

    ' Position primitives
    Face.ObjRotY = headAngle
    FaceGuides.ObjRotY = headAngle
    pFaceDiverter.ObjRotY = headAngle
    pFaceDiverterPegs.ObjRotY = headAngle


  ' Determine which face is up, if any
   If headIsNear(0) OR headIsNear(360) Then
      ShowEyeBulbs(True)
      currentFace = 1

	If HottieModType = 1 Then

	Else
		'FaceGuides.Visible = true
	End If
 l46_b.State = 1
 l47_b.State = 1
 l46_face2.State = 0
 l47_face2.State = 0
      DOF 102, DOFOff
    ElseIf headIsNear(90) Then
      ShowEyeBulbs(True)
      currentFace = 2
	If HottieModType = 1 Then

	Else
		FaceGuides.Visible = true
	End If
 l46_b.State = 0
 l47_b.State = 0
 l46_face2.State = 1
 l47_face2.State = 1
      DOF 102, DOFoff
    ElseIf headIsNear(180) Then
      ShowEyeBulbs(True)
      currentFace = 3
	If HottieModType = 1 Then

	Else
		FaceGuides.Visible = true
	End If
 l46_b.State = 1
 l47_b.State = 1
 l46_face2.State = 0
 l47_face2.State = 0
      DOF 102, DOFOff
    ElseIf headIsNear(270) Then
      ShowEyeBulbs(True)
      currentFace = 4
	If HottieModType = 1 Then

	Else
		FaceGuides.Visible = True
	End If
 l46_b.State = 1
 l47_b.State = 1
 l46_face2.State = 0
 l47_face2.State = 0
      DOF 102, DOFOff
    Else
      ShowEyeBulbs(False)
      currentFace = 0
      DOF 102, DOFOn
    End If

    ' Head box position switch
   Controller.Switch(67) = (currentFace > 0 AND currentFace < 4) ' Faces 1, 2 and 3 close the switch.

    ' Face 1 parts
   sw65.Enabled = (currentFace = 1)
    F1Guide.IsDropped = NOT (currentFace = 1)
    F1Guide2.IsDropped = NOT (currentFace = 1)

    ' Face 2 parts
   sw63.Enabled = (currentFace = 2)
    sw64.Enabled = (currentFace = 2)
    div.Enabled = (currentFace = 2)
    sw63div.Enabled = (currentFace = 2)
    sw64div.Enabled = (currentFace = 2)
End Sub

Sub ShowEyeBulbs(vis)

      h23.visible = vis
      h42.visible = vis
      element18.visible = vis
      element26.visible = vis
      wall30.isdropped = not(vis)
      wall34.isdropped = not(vis)
      wall35.isdropped = not(vis)
      wall38.isdropped = not(vis)
      wall39.isdropped = not(vis)
      wall87.isdropped = not(vis)

      If vis Then
            L46_a.intensity = 4
            L47_a.intensity = 4

      Else
            L46_a.intensity = 0
            L47_a.intensity = 0

      End If


End Sub

'-----------------
' End Head/Face rotation script
'-----------------

 'Kickers, poppers
  Sub sw63_Hit():bsLEye.AddBall 0: End Sub
   Sub sw64_Hit():bsREye.AddBall 0: End Sub

   Sub sw63div_hit():div.RotateToStart:End Sub
   Sub sw64div_Hit():div.RotateToEnd:End Sub

   Sub sw65_Hit():bsMouth.AddBall 0:End Sub
   Sub SSLaunch_Hit():activeball.mass = BallMass:bsSS.AddBall 0:End Sub

   'Bumpers
     Sub Bumper1b_Hit:vpmTimer.PulseSw 53:PlaySound SoundFx("fx_bumper_1",DOFContactors):End Sub 'bump1 = 1:Me.TimerEnabled = 1

       Sub Bumper1b_Timer()  '53
          Select Case bump1
               Case 1:BR1.z = 0:bump1 = 2
               Case 2:BR1.z = -10:bump1 = 3
               Case 3:BR1.z = -20:bump1 = 4
               Case 4:BR1.z = -20:bump1 = 5
               Case 5:BR1.z = -10:bump1 = 6
               Case 6:BR1.z = 0:bump1 = 7
               Case 7:BR1.z = 10:Me.TimerEnabled = 0
           End Select
       End Sub

      Sub Bumper2b_Hit:vpmTimer.PulseSw 55:PlaySound SoundFx("fx_bumper_2",DOFContactors):End Sub 'bump2 = 1:Me.TimerEnabled = 1

       Sub Bumper2b_Timer()   '55
          Select Case bump2
               Case 1:BR2.z = 0:bump2 = 2
               Case 2:BR2.z = -10:bump2 = 3
               Case 3:BR2.z = -20:bump2 = 4
               Case 4:BR2.z = -20:bump2 = 5
               Case 5:BR2.z = -10:bump2 = 6
               Case 6:BR2.z = 0:bump2 = 7
               Case 7:BR2.z = 10:Me.TimerEnabled = 0
           End Select
       End Sub

      Sub Bumper3b_Hit:vpmTimer.PulseSw 54:PlaySound SoundFx("fx_bumper_3",DOFContactors):End Sub 'bump3 = 1:Me.TimerEnabled = 1

       Sub Bumper3b_Timer()  '54
          Select Case bump3
               Case 1:BR3.z = 0:bump3 = 2
               Case 2:BR3.z = -10:bump3 = 3
               Case 3:BR3.z = -20:bump3 = 4
               Case 4:BR3.z = -20:bump3 = 5
               Case 5:BR3.z = -10:bump3 = 6
               Case 6:BR3.z = 0:bump3 = 7
               Case 7:BR3.z = 10:Me.TimerEnabled = 0
           End Select

       End Sub

 'StandUp Targets
  Sub sw28_Hit:vpmTimer.PulseSw 28:Me.TimerEnabled = 1:PlaySound SoundFx("target",DOFTargets):End Sub
   Sub sw28_Timer:Me.TimerEnabled = 0:End Sub

   Sub sw36_Hit:vpmTimer.PulseSw 36:Me.TimerEnabled = 1:PlaySound SoundFx("target",DOFTargets):End Sub
   Sub sw36_Timer:Me.TimerEnabled = 0:End Sub

   Sub sw37_Hit:vpmTimer.PulseSw 37:Me.TimerEnabled = 1:PlaySound SoundFx("target",DOFTargets):End Sub
   Sub sw37_Timer:Me.TimerEnabled = 0:End Sub


 'FlipperLanes and Plunger
  Sub sw15_Hit:Controller.Switch(15) = 1:PlaySound "rollover":End Sub
   Sub sw15_UnHit:Controller.Switch(15) = 0:End Sub
   Sub sw16_Hit:Controller.Switch(16) = 1:PlaySound "rollover":End Sub
   Sub sw16_UnHit:Controller.Switch(16) = 0:End Sub
   Sub sw17_Hit:Controller.Switch(17) = 1:PlaySound "rollover":End Sub
   Sub sw17_UnHit:Controller.Switch(17) = 0:End Sub
   Sub sw18_Hit:Controller.Switch(18) = 1:PlaySound "rollover":End Sub
   Sub sw18_UnHit:Controller.Switch(18) = 0:End Sub
   Sub sw52_Hit:Controller.Switch(52) = 1:activeball.mass = 1:PlaySound "rollover":End Sub
   Sub sw52_UnHit:Controller.Switch(52) = 0:End Sub

   'SS Lane
  Sub sw31_Hit:Controller.Switch(31) = 1:PlaySound "rollover":End Sub
   Sub sw31_UnHit:Controller.Switch(31) = 0:End Sub
   Sub sw32_Hit:Controller.Switch(32) = 1:PlaySound "rollover":End Sub
   Sub sw32_UnHit:Controller.Switch(32) = 0:End Sub
   Sub sw33_Hit:Controller.Switch(33) = 1:PlaySound "rollover":End Sub
   Sub sw33_UnHit:Controller.Switch(33) = 0:End Sub
   Sub sw34_Hit:Controller.Switch(34) = 1:PlaySound "rollover":End Sub
   Sub sw34_UnHit:Controller.Switch(34) = 0:End Sub
   Sub sw35_Hit:Controller.Switch(35) = 1:PlaySound "rollover":End Sub

   Sub sw35_UnHit:Controller.Switch(35) = 0:End Sub
   'Under Right Ramp
  Sub sw44_Hit:Controller.Switch(44) = 1:PlaySound "rollover":End Sub
   Sub sw44_UnHit:Controller.Switch(44) = 0:End Sub
   Sub sw45_Hit:Controller.Switch(45) = 1:PlaySound "rollover":End Sub
   Sub sw45_UnHit:Controller.Switch(45) = 0:End Sub

   'Left Loop
  Sub sw43_Hit:Controller.Switch(43) = 1:PlaySound "rollover":End Sub
   Sub sw43_UnHit:Controller.Switch(43) = 0:End Sub

   'Gate Switches
  Sub sw47_Hit:vpmTimer.PulseSw 47:End Sub

   'Ball lock Switches
    'LockInit
  LockWall.IsDropped = 1
   Sub sw72_Hit:Controller.Switch(72) = 1:::PlaySound "rollover":LockWall.IsDropped = 0:Switch72dir=-2:sw72.TimerEnabled = True:End Sub
   Sub sw72_UnHit:Controller.Switch(72) = 0:Switch72dir=4:sw72.TimerEnabled = True::LockWall.IsDropped = 1:PlaySound "rail_low_slower":End Sub
   Sub sw71_Hit:Controller.Switch(71) = 1::PlaySound "rollover":Switch71dir=-2:sw71.TimerEnabled = True:End Sub
   Sub sw71_UnHit:Controller.Switch(71) = 0:Switch71dir=4:sw71.TimerEnabled = True:End Sub

Sub sw41_Hit:vpmTimer.PulseSw(41):sw41.TimerEnabled = True:End Sub
Sub sw73_Hit:vpmTimer.PulseSw(73):sw73.TimerEnabled = True:End Sub
Sub sw74_Hit:vpmTimer.PulseSw(74):sw74.TimerEnabled = True:PlaySound "fx_Gate":End Sub
Sub sw75_Hit:vpmTimer.PulseSw(75):sw75.TimerEnabled = True:End Sub
Sub sw76_Hit:vpmTimer.PulseSw(76):sw76.TimerEnabled = True:PlaySound "fx_Gate":End Sub
Sub sw77_Hit:vpmTimer.PulseSw(77):sw77.TimerEnabled = True:PlaySound "fx_Gate":End Sub

Const Switch41min = 0
Const Switch41max = -20
Dim Switch41dir

    'switch 41 animation
Switch41dir = -2

Sub sw41_timer()
 pRampSwitch1B.RotY = pRampSwitch1B.RotY + Switch41dir
    If pRampSwitch1B.RotY >= Switch41min Then
        sw41.timerenabled = False
        pRampSwitch1B.RotY = Switch41min
        Switch41dir = -2
    End If
    If pRampSwitch1B.RotY <= Switch41max Then
        Switch41dir = 4
    End If
End Sub

    'switch 71 animation
Const Switch71min = 0
Const Switch71max = -20
Dim Switch71dir
Switch71dir = -2

Sub sw71_timer()
 pRampSwitch7B.RotY = pRampSwitch7B.RotY + Switch71dir
    If Switch71dir = 4 Then
        If pRampSwitch7B.RotY >= Switch71min Then
            sw71.timerenabled = False
            pRampSwitch7B.RotY = Switch71min
        End If
    End If
    If Switch71dir = -2 Then
        If pRampSwitch7B.RotY <= Switch71max Then
            sw71.timerenabled = False
            pRampSwitch7B.RotY = Switch71max
        End If
    End If
End Sub

    'switch 72 animation
Const Switch72min = 0
Const Switch72max = -20
Dim Switch72dir
Switch72dir = -2

Sub sw72_timer()
 pRampSwitch8B.RotY = pRampSwitch8B.RotY + Switch72dir
    If pRampSwitch8B.RotY >= Switch72min Then
        sw72.timerenabled = False
        pRampSwitch8B.RotY = Switch72min
    End If
    If pRampSwitch8B.RotY <= Switch72max Then
        sw72.timerenabled = False
        pRampSwitch8B.RotY = Switch72max
    End If
End Sub

    'switch 73 animation
Const Switch73min = 0
Const Switch73max = -20
Dim Switch73dir
Switch73dir = -2

Sub sw73_timer()
 pRampSwitch6B.RotX = pRampSwitch6B.RotX + Switch73dir
    If pRampSwitch6B.RotX >= Switch73min Then
        sw73.timerenabled = False
        pRampSwitch6B.RotX = Switch73min
        Switch73dir = -2
    End If
    If pRampSwitch6B.RotX <= Switch73max Then
        Switch73dir = 4
    End If
End Sub

    'switch 74 animation
Const Switch74min = 0
Const Switch74max = -20
Dim Switch74dir
Switch74dir = -2

Sub sw74_timer()
 pRampSwitch3B.RotX = pRampSwitch3B.RotX + Switch74dir
    If pRampSwitch3B.RotX >= Switch74min Then
        sw74.timerenabled = False
        pRampSwitch3B.RotX = Switch74min
        Switch74dir = -2
    End If
    If pRampSwitch3B.RotX <= Switch74max Then
        Switch74dir = 4
    End If
End Sub

    'switch 75 animation
Const Switch75min = 0
Const Switch75max = -20
Dim Switch75dir
Switch75dir = -2

Sub sw75_timer()
 pRampSwitch2B.RotY = pRampSwitch2B.RotY + Switch75dir
    If pRampSwitch2B.RotY >= Switch75min Then
        sw75.timerenabled = False
        pRampSwitch2B.RotY = Switch75min
        Switch75dir = -2
    End If
    If pRampSwitch2B.RotY <= Switch75max Then
        Switch75dir = 4
    End If
End Sub

    'switch 76 animation
Const Switch76min = 0
Const Switch76max = -20
Dim Switch76dir
Switch76dir = -2

Sub sw76_timer()
 pRampSwitch5B.RotX = pRampSwitch5B.RotX + Switch76dir
    If pRampSwitch5B.RotX >= Switch76min Then
        sw76.timerenabled = False
        pRampSwitch5B.RotX = Switch76min
        Switch76dir = -2
    End If
    If pRampSwitch5B.RotX <= Switch76max Then
        Switch76dir = 4
    End If
End Sub

    'switch 77 animation
Const Switch77min = 0
Const Switch77max = -20
Dim Switch77dir
Switch77dir = -2

Sub sw77_timer()
 pRampSwitch4B.Rotx = pRampSwitch4B.Rotx + Switch77dir
    If pRampSwitch4B.Rotx >= Switch77min Then
        sw77.timerenabled = False
        pRampSwitch4B.Rotx = Switch77min
        Switch77dir = -2
    End If
    If pRampSwitch4B.Rotx <= Switch77max Then
        Switch77dir = 4
    End If
End Sub


'*****************
'Animated rubbers
'*****************
Sub wall69_Hit:vpmTimer.PulseSw 100:rubber25.visible = 0::rubber25a.visible = 1:wall69.timerenabled = 1:End Sub
Sub wall69_timer:rubber25.visible = 1::rubber25a.visible = 0: wall69.timerenabled= 0:End Sub

Sub wall75_Hit:vpmTimer.PulseSw 100:rubber24.visible = 0::rubber24a.visible = 1:wall75.timerenabled = 1:End Sub
Sub wall75_timer:rubber24.visible = 1::rubber24a.visible = 0: wall75.timerenabled= 0:End Sub

Sub wall77_Hit:vpmTimer.PulseSw 100:rubber23.visible = 0::rubber23a.visible = 1:wall77.timerenabled = 1:End Sub
Sub wall77_timer:rubber23.visible = 1::rubber23a.visible = 0: wall77.timerenabled= 0:End Sub

Sub wall79_Hit:vpmTimer.PulseSw 100:rubber22.visible = 0::rubber22a.visible = 1:wall79.timerenabled = 1:End Sub
Sub wall79_timer:rubber22.visible = 1::rubber22a.visible = 0: wall79.timerenabled= 0:End Sub

Sub wall80_Hit:vpmTimer.PulseSw 100:rubber21.visible = 0::rubber21a.visible = 1:wall80.timerenabled = 1:End Sub
Sub wall80_timer:rubber21.visible = 1::rubber21a.visible = 0: wall80.timerenabled= 0:End Sub

Sub wall81_Hit:vpmTimer.PulseSw 100:rubber20.visible = 0::rubber20a.visible = 1:wall81.timerenabled = 1:End Sub
Sub wall81_timer:rubber20.visible = 1::rubber20a.visible = 0: wall81.timerenabled= 0:End Sub

Sub wall82_Hit:vpmTimer.PulseSw 100:rubber3.visible = 0::rubber3a.visible = 1:wall82.timerenabled = 1:End Sub
Sub wall82_timer:rubber3.visible = 1::rubber3a.visible = 0: wall82.timerenabled= 0:End Sub

Sub wall83_Hit:vpmTimer.PulseSw 100:rubber14.visible = 0::rubber14a.visible = 1:wall83.timerenabled = 1:End Sub
Sub wall83_timer:rubber14.visible = 1::rubber14a.visible = 0: wall83.timerenabled= 0:End Sub

Sub wall84_Hit:vpmTimer.PulseSw 100:rubber19.visible = 0::rubber19a.visible = 1:wall84.timerenabled = 1:End Sub
Sub wall84_timer:rubber19.visible = 1::rubber19a.visible = 0: wall84.timerenabled= 0:End Sub

Sub LockWall_Hit():Playsound "collide0":End Sub

'Spinner
Sub sw51_Spin:vpmTimer.PulseSw 51:PlaySound "spinner":End Sub


'Ball Drop
Sub RRail_Hit()
    StopRollingSound
    dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
    if finalspeed > 2 AND finalspeed < 4 Then
        PlaySound "rail_low_slower",0,1,0,0
    elseif finalspeed >= 4 Then
        PlaySound "rail",0,1,0,0
    End if
End Sub




Sub LRail_Hit()
    StopRollingSound
    dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
    if finalspeed > 2 AND finalspeed < 4 Then
        PlaySound "rail_low_slower",0,1,0,0
    elseif finalspeed >= 4 Then
        PlaySound "rail",0,1,0,0
    End if
End Sub

Sub UpperRubbersBandsLargeRings_Hit(idx)
    dim finalspeed
    finalspeed = SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
    If finalspeed > 20 then
        PlaySound "fx_rubber", 0, 3*Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    End if
    If finalspeed >= 6 AND finalspeed <= 20 then
        RandomSoundRubber()
    End If
End Sub

Sub UpperRubbersSmallRings_Hit(idx)
    dim finalspeed
    finalspeed = SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
    If finalspeed > 20 then
        PlaySound "fx_rubber", 0, 3*Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    End if
    If finalspeed >= 6 AND finalspeed <= 20 then
        RandomSoundRubber()
    End If
End Sub

Sub UpperRubbersWalls_Hit(idx)
    dim finalspeed
    finalspeed = SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
    If finalspeed > 20 then
        PlaySound "fx_rubber", 0, 3*Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    End if
    If finalspeed >= 6 AND finalspeed <= 20 then
        RandomSoundRubber()
    End If
End Sub

Sub Skillrubbers_Hit(idx)
    dim finalspeed
    finalspeed = SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
    If finalspeed > 20 then
        PlaySound "fx_rubber", 0, 3*Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    End if
    If finalspeed >= 6 AND finalspeed <= 20 then
        RandomSoundRubber()
    End If
End Sub

Sub RubbersSmallRings_Hit(idx)
    dim finalspeed
    finalspeed = SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
    If finalspeed > 20 then
        PlaySound "fx_rubber", 0, 3*Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    End if
    If finalspeed >= 6 AND finalspeed <= 20 then
        RandomSoundRubber()
    End If
End Sub

Sub LowerRubbersBandsLargeRings_Hit(idx)
    dim finalspeed
    finalspeed = SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
    If finalspeed > 20 then
        PlaySound "fx_rubber", 0, 3*Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    End if
    If finalspeed >= 6 AND finalspeed <= 20 then
        RandomSoundRubber()
    End If
End Sub

Sub RandomSoundRubber()
    Select Case Int(Rnd * 3) + 1
        Case 1:PlaySound "fx_rubber_hit_1", 0, 2*Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
        Case 2:PlaySound "fx_rubber_hit_2", 0, 2*Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
        Case 3:PlaySound "fx_rubber_hit_3", 0, 2*Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
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
		Case 1 : PlaySound "fx_flip_hit_1", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 2 : PlaySound "fx_flip_hit_2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 3 : PlaySound "fx_flip_hit_3", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End Select
End Sub

'*****************
' GI2 Illumination (Variable Lights 1-8 and off - Williams / Bally)
' 8 Step Sub
'*****************

'***********
' Update GI
'***********
'*****Init gi

REM Dim gistep
REM gistep = 1 / 8
REM UpdateGI 0,0:UpdateGI 1,0:UpdateGI 2,0:UpdateGI 3,0:UpdateGI 4,0
REM Sub UpdateGI(no, step)
    REM If step = 7 then exit sub    '0 OR step =
   REM Select Case no
        REM Case 0  'backglass
REM ''            For each xx in BlueGIHalo:xx.alpha = gistep * step:next
REM ''            GISL.alpha = gistep * step

        REM Case 1   'helmet
REM '           HelGI.IntensityScale = cint(gistep * step)
REM '           HelGI1.IntensityScale = cint(gistep * step)
       REM Case 2   'rear pf
            REM For each xx in GI_rear:xx.intensityscale = gistep * step:next
                 REM 'Playsound "fx_relay_on"
       REM Case 3   'backglass
REM '
       REM Case 4  'front pf
            REM For each xx in GI_front:xx.IntensityScale = gistep * step:next
                 REM 'Playsound "fx_relay_on"
   REM End Select

    REM Table.ColorGradeImage = "ColorGradeBOP_" & step

REM End Sub

'================Light Handling==================
'       Gi, Flashers, and Lamp handling
'Based on JP's VP10 fading Lamp routine,  based on PD's Fading Lights
'       Mod FrameTime and GI handling by nFozzy
'================================================

Dim LampState(340), FadingLevel(340), CollapseMe
Dim FlashSpeedUp(340), FlashSpeedDown(340), FlashMin(340), FlashMax(340), FlashLevel(340)
dim SolModValue(340)    'holds 0-255 modulated solenoid values

'These are used for fading lights and flashers brighter when the GI is darker
DIm LampsOpacity(340, 2) 'Columns: 0 = intensity / opacity, 1 = fadeup, 2 = FadeDown
dim GIscale(3)  '4 gi strings

InitLamps

redim CollapseMe(0) 'Setlamps and SolModCallBacks   (Click Me to Collapse)
    Sub SetLamp(nr, value)
        If value <> LampState(nr) Then
            LampState(nr) = abs(value)
            FadingLevel(nr) = abs(value) + 4
        End If
    End Sub

    Sub SetLampm(nr, nr2, value)    'set 2 lamps
        If value <> LampState(nr) Then
            LampState(nr) = abs(value)
            FadingLevel(nr) = abs(value) + 4
        End If
        If value <> LampState(nr2) Then
            LampState(nr2) = abs(value)
            FadingLevel(nr2) = abs(value) + 4
        End If
    End Sub

    Sub SetModLamp(nr, value)
        If value <> SolModValue(nr) Then
            SolModValue(nr) = value
            if value > 0 then LampState(nr) = 1 else LampState(nr) = 0
            FadingLevel(nr) = LampState(nr) + 4
        End If
    End Sub

    Sub SetModLampM(nr, nr2, value) 'set 2 modulated lamps
        If value <> SolModValue(nr) Then
            SolModValue(nr) = value
            if value > 0 then LampState(nr) = 1 else LampState(nr) = 0
            FadingLevel(nr) = LampState(nr) + 4
        End If
        If value <> SolModValue(nr2) Then
            SolModValue(nr2) = value
            if value > 0 then LampState(nr2) = 1 else LampState(nr2) = 0
            FadingLevel(nr2) = LampState(nr2) + 4
        End If
    End Sub
    'Flashers via SolModCallBacks
  SolModCallBack(17) = "SetModLamp 117," 'Billions
  SolModCallBack(18) = "SetModLamp 118," 'Left ramp
  SolModCallBack(19) = "SetModLamp 119," 'jackpot
  SolModCallBack(20) = "SetModLamp 120," 'SkillShot
  SolModCallBack(21) = "SetDome1"'"SetModLamp 121," 'Left Helmet
  SolModCallBack(22) = "SetDome2"'"SetModLamp 122," 'Right Helmet
  SolModCallBack(23) = "SetDome3"'"SetModLamp 123," 'Jets Enter
  SolModCallBack(24) = "SetDome4"'"SetModLamp 124," 'Left Loop

    Sub SetDome1(value)     ' *** Helmet Left Flasher ***
        SetModLamp 121, value
        if value Then
            Flasher1.image="domeon1"
            Flasher1.material="flasheron"
            Flasher1.disablelighting=1
        Else
            Flasher1.image="domeoff"
            Flasher1.material="flasheroff"
            Flasher1.disablelighting=0
        End If
    End Sub

    Sub SetDome2(value)     ' *** Helmet Right Flasher ***
        SetModLamp 122, value
        if value Then
            Flasher2.image="domeon1"
            Flasher2.material="flasheron"
            Flasher2.disablelighting=1
        Else
            Flasher2.image="domeoff"
            Flasher2.material="flasheroff"
            Flasher2.disablelighting=0
        End If
    End Sub

    Sub SetDome3(value)     ' *** Right Boob Flasher ***
        SetModLamp 123, value
        if value Then
            Flasher3.image="domeon1"
            Flasher3.material="flasheron"
            Flasher3.disablelighting=1
        Else
            Flasher3.image="domeoff"
            Flasher3.material="flasheroff"
            Flasher3.disablelighting=0
        End If
    End Sub

    Sub SetDome4(value)     ' *** Left Boob Flasher ***
        SetModLamp 124, value
        if value Then
            Flasher4.image="domeon1"
            Flasher4.material="flasheron"
            Flasher4.disablelighting=1
        Else
            Flasher4.image="domeoff"
            Flasher4.material="flasheroff"
            Flasher4.disablelighting=0
        End If
    End Sub

'-------------------------------------
redim CollapseMe(0) 'InitLamps  (Click Me to Collapse)
    Sub InitLamps() 'set fading speeds and other stuff here
        GetOpacity aLampsAll    'All non-GI lamps and flashers go in this object array for compensation script!
        dim x
        for x = 0 to uBound(LampState)
            On Error Resume Next
            LampState(x) = 0    ' current light state, independent of the fading level. 0 is off and 1 is on
            FadingLevel(x) = 4  ' used to track the fading state
            FlashSpeedUp(x) = 0.1   'Fading speeds in opacity per MS I think (Not used with nFadeL or nFadeLM subs!)
            FlashSpeedDown(x) = 0.1

            FlashMin(x) = 0.001         ' the minimum value when off, usually 0
            FlashMax(x) = 1             ' the minimum value when off, usually 1
            FlashLevel(x) = 0.001       ' Raw Flasher opacity value. Start this >0 to avoid initial flasher stuttering.

            SolModValue(x) = 0          ' Holds SolModCallback values

            Giscale(x) = 1.625          ' lamp GI compensation multiplier, eg opacity x 1.625 when gi is fully off
        Next

        for x = 11 to 110 'insert fading levels (only applicable for lamps that use FlashC sub)
            FlashSpeedUp(x) = 0.015
            FlashSpeedDown(x) = 0.009
        Next

        for x = 111 to 186  'Flasher fading speeds 'intensityscale(%) per 10MS
            FlashSpeedUp(x) = 1.1
            FlashSpeedDown(x) = 0.9
        next

        for x = 200 to 203      'GI relay on / off  fading speeds
            FlashSpeedUp(x) = 0.01
            FlashSpeedDown(x) = 0.008
            FlashMin(x) = 0
        Next
        for x = 300 to 303      'GI 8 step modulation fading speeds
            FlashSpeedUp(x) = 0.01
            FlashSpeedDown(x) = 0.008
            FlashMin(x) = 0
        Next

        UpdateGIon 0, 1:UpdateGIon 1, 1: UpdateGIon 2, 1: UpdateGIon 3, 1: UpdateGIon 4, 1
        UpdateGI 0, 7:UpdateGI 1, 7:UpdateGI 2, 7: UpdateGI 3, 7: UpdateGI 4, 7
    End Sub

    Sub GetOpacity(a)   'Keep lamp/flasher data in an array
        dim x
        for x = 0 to (a.Count - 1)
            On Error Resume Next
            if a(x).Opacity > 0 then a(x).Uservalue = a(x).Opacity
            if a(x).Intensity > 0 then a(x).Uservalue = a(x).Intensity
            If a(x).FadeSpeedUp > 0 then LampsOpacity(x, 1) = a(x).FadeSpeedUp : LampsOpacity(x, 2) = a(x).FadeSpeedDown
        Next
        for x = 0 to (a.Count - 1)
            LampsOpacity(x, 0) = a(x).UserValue

            REM if a(x).name = "Flasher124" then
                REM tb.text = a(x).name & vbnewline & _
                 REM a(x).Opacity & vbnewline & _
                 REM a(x).UserValue & vbnewline & _
                REM LampsOpacity(x,0)
            REM end if
        Next

        for each x in a : x.state = 1 : Next
        for each x in GI_Rear : x.state = 1 : Next
        for each x in GI_Front : x.state = 1 : Next

    End Sub

    sub DebugLampsOn(input):dim x: for x = 10 to 100 : setlamp x, input : next :  end sub

'----------------------------------

redim CollapseMe(0) 'LampTimer  (Click Me to Collapse)

    LampTimer.Interval = -1 '-1 is ideal, but it will technically work with any timer interval
    dim FrameTime, InitFadeTime : FrameTime = 10    'Count Frametime
    Sub LampTimer_Timer()
        FrameTime = gametime - InitFadeTime
        Dim chgLamp, num, chg, ii
        chgLamp = Controller.ChangedLamps
        If Not IsEmpty(chgLamp) Then
            For ii = 0 To UBound(chgLamp)
                LampState(chgLamp(ii, 0) ) = chgLamp(ii, 1)       'keep the real state in an array
                FadingLevel(chgLamp(ii, 0) ) = chgLamp(ii, 1) + 4 'actual fading step
            Next
        End If

        UpdateGIstuff
        UpdateLamps
        UpdateFlashers

        InitFadeTime = gametime
    End Sub
'#end section
redim CollapseMe(0) 'ASSIGNMENTS: Lamps, GI, and Flashers (Click Me to Collapse)
    Sub UpdateGIstuff()
        FadeGI 200  '200/300 Backglass #1
        ModGI  300
        FadeGI 201  '201/301 Helmet (?)
        ModGI  301
        FadeGI 202  '202/302 Rear PF
        ModGI  302
        FadeGI 203  '203/303 Backglass #2
        ModGI  303
        FadeGI 204  '204/304 Front PF
        ModGI  304
        UpdateGIObjects 202, 302, GI_Rear   'nr, nr2, array 'Updates GI objects
        UpdateGIObjects 204, 304, GI_Front  'nr, nr2, array 'Updates GI objects
        'GIcompensation 202, 302, LampsRear, GiScale(2) 'nr, nr2, Lamp Array to Tweak, GiScaler mult
        GiCompensationAvg 202, 302, 204, 304, aLampsAll, GiScale(2) 'Averages two lamp strings. Ideal match the avg. Lut fading.
        'FadeLut 202, 302, "LutSF_", 21 'Nr1(on/off), Nr2(Mod), LUT name prefix, max number of luts
        FadeLutAvg 202, 302, 204, 304, "LutCont_", 27   'Lut averages two fading strings
    End Sub

    Sub UpdateFlashers()
        nModFlash 118, f18, 0, 0    'NR, Object, LowPass(see function ScaleByte) OffScale (multiply fading when callback 0)

        nModFlash 117, f17, 0, 0    'Last two arguments will carry through to objects controlled by nModFlashM.
        nModFlash 119, f19, 0, 0    'For more control, use multiple fading numbers via SetModLampM / SetModLampMM
        nModFlash 120, L120, 0, 0

        nModFlash 121, Flasher121, 0, 0
        nModFlash 122, Flasher122, 0, 0
        nModFlash 123, Flasher123, 0, 0
        nModFlash 124, Flasher124, 0, 0

        If ShipMod = 1 then
            nModFlashm 124, F124c
            nModFlashm 124, F124d
            nModFlashm 124, F124e
            nModFlashm 124, F124f
            nModFlashm 124, F124g
            nModFlashm 124, F124h
        End If
    End Sub

    Sub UpdateLamps()
	If LampState(11) = 1 Then
		If MusicSnippet = 0 And PrevGameOver = 0 Then
			'StopSound "intro"
			PlaySound "intro"
			PrevGameOver = 1
		End If
	else
'		PrevGameOver = 0
	End If

        nFadeL 11, l11
        nFadeL 12, l12
        nFadeL 13, l13
        nFadeL 14, l14
        nFadeL 15, l15
        nFadeL 16, l16
        nFadeL 17, l17
        nFadeL 18, l18
        nFadeL 21, l21
        nFadeL 22, l22
        nFadeL 23, l23
        nFadeL 24, l24
        nFadeL 25, l25
        nFadeL 26, l26
        nFadeL 27, l27
        nFadeL 28, l28
        nFadeL 31, l31
        nFadeL 32, l32
        nFadeL 33, l33
        nFadeL 34, l34
        nFadeL 35, l35
        nFadeL 36, l36
        nFadeLm 37, l37
        nFadeLm 37, l37a
        FlashC 37, f37
        nFadeLm 38, l38a
        nFadeLm 38, l38b
        nFadeLm 38, l38c
        nFadeLm 38, l38d
        nFadeLm 38, l38e
        nFadeLm 38, l38f
        nFadeLm 38, l38g
        nFadeLm 38, l38h
        nFadeLm 38, l38i
        nFadeLm 38, l38j
        nFadeLm 38, l38k
        nFadeLm 38, l38l
        nFadeLm 38, l38m
        nFadeL 38, l38

        nFadeLFOm 41
        nFadePrim 41, p50k, "ScoreDecals50kon_a", "ScoreDecals"
		If SSRampModType = 2 Then nFadePrim 41, Primitive53, "SkillShotRamp_19", "SkillShotRamp_allOff" End If
        nFadelm 41, Light19b
        nFadelm 41, Light19b2
        nFadel 41, Light19b1

        nFadeLFOm 42
        nFadePrim 42, p75k, "ScoreDecals75kon_a", "ScoreDecals"
		If SSRampModType = 2 Then nFadePrim 42, Primitive53, "SkillShotRamp_20", "SkillShotRamp_allOff" End If
        nFadelm 42, Light20b
        nFadelm 42, Light20b2
        nFadel 42, Light20b1

        nFadeLFOm 43
        nFadePrim 43, p100k, "ScoreDecals100kon_a", "ScoreDecals"
		If SSRampModType = 2 Then nFadePrim 43, Primitive53, "SkillShotRamp_21", "SkillShotRamp_allOff" End If
        nFadelm 43, Light21b
        nFadelm 43, Light21b2
        nFadel 43, Light21b1

        nFadeLFOm 44
        nFadePrim 44, p200k, "ScoreDecals200kon_a", "ScoreDecals"
		If SSRampModType = 2 Then nFadePrim 44, Primitive53, "SkillShotRamp_22", "SkillShotRamp_allOff" End If
        nFadelm 44, Light22b
        nFadelm 44, Light22b2
        nFadel 44, Light22b1

        nFadeLFOm 45
        nFadePrim 45, p25k, "ScoreDecals25kon_a", "ScoreDecals"
		If SSRampModType = 2 Then nFadePrim 45, Primitive53, "SkillShotRamp_23", "SkillShotRamp_allOff" End If
        nFadeLm 45, Light23b
        nFadeLm 45, Light23b2
        nFadeL 45, Light23b1

		If SSRampModType = 2 Then
			If LampState(41) = 1 and LampState(42) = 1 and LampState(43) = 1 and LampState(44) = 1 and LampState(45) = 1 Then
				'Primitive68.image = "scoreplastic_allOn"
				Primitive53.image = "SkillShotRamp_allOn"
			End If
		End If

        nFadeLm 46, l46_a
        nFadeLm 46, l46_c
		if currentFace = 2 then
			nFadeLm 46, l46_face2
		Else
			l46_face2.state = 0
		End If
        nFadeL 46, l46_b
        nFadeLm 47, l47_a
        nFadeLm 47, l47_c
		if currentFace = 2 then
			nFadeLm 46, l47_face2
		Else
			l47_face2.state = 0
		End If
        nFadeL 47, l47_b
        nFadelm 48, l48_a
        nFadelm 48, l48_c
        nFadel 48, l48_b
        nFadeL 51, l51
        nFadeL 52, l52
        nFadeL 53, l53
        nFadeL 54, l54
        nFadeLm 55, l55a
        nFadeL 55, l55
        nFadeL 56, l56
        nFadeL 57, l57
        nFadeL 58, l58
        nFadeL 61, l61
        nFadeL 62, l62
        nFadeL 63, l63
        nFadeLFOm 64
        Flashc 64, F64a
        Flashm 64, F64b
        Flashc 86, f86
        Flashc 87, f87
        Flashc 88, f88
        nFadeL 65, l65
        nFadeL 66, l66
        nFadeL 67, l67
        nFadeL 68, l68

        'Helmet Lights
        nFadeLm 91, l91
        FlashC 91, f91

        nFadeLm 92, l92
        FlashC 92, f92

        nFadeLm 93, l93
        FlashC 93, f93

        nFadeLm 94, l94
        FlashC 94, f94

        nFadeLm 95, l95
        FlashC 95, f95

        nFadeLm 96, l96
        FlashC 96, f96

        nFadeLm 97, l97
        FlashC 97, f97

        nFadeLm 98, l98
        FlashC 98, f98

        nFadeLm 101, l101
        FlashC 101, f101

        nFadeLm 102, l102
        FlashC 102, f102

        nFadeLm 103, l103
        FlashC 103, f103

        nFadeLm 104, l104
        FlashC 104, f104

        nFadeLm 105, l105
        FlashC 105, f105

        nFadeLm 106, l106
        FlashC 106, f106

        nFadeLm 107, l107
        FlashC 107, f107

        nFadeLm 108, l108
        FlashC 108, f108


		If GIColorModType = 2 Then
			nFadePrim 234, Primitive4, "leftramp_blue", "leftramp 0.28"
			nFadePrim 234, Primitive47, "heartramp_Blue_Upper", "heartramp 0.28b test"
			nFadePrim 235, Primitive47, "heartramp_Blue_Lower", "heartramp 0.28b test"
			If LampState(234) = 1 and LampState(235) = 1Then
				Primitive47.image = "heartramp_Blue_AllOn"
			End If
			If LampState(204) = 0 Then
				nFadePrim 235, Primitive47, "heartramp 0.28b test", "heartramp 0.28b test"
				nFadePrim 235, Primitive4, "leftramp 0.28", "leftramp 0.28"
			End If
		End If

		If GIColorModType = 3 Then
			nFadePrim 234, Primitive4, "leftramp_pink", "leftramp 0.28"
			nFadePrim 234, Primitive47, "heartramp_Pink_Upper", "heartramp 0.28b test"
			nFadePrim 235, Primitive47, "heartramp_Pink_Lower", "heartramp 0.28b test"
			If LampState(234) = 1 and LampState(235) = 1Then
				Primitive47.image = "heartramp_Pink_AllOn"
			End If
			If LampState(204) = 0 Then
				nFadePrim 235, Primitive47, "heartramp 0.28b test", "heartramp 0.28b test"
				nFadePrim 235, Primitive4, "leftramp 0.28", "leftramp 0.28"
			End If
		End If
		If GIColorModType = 4 Then
			nFadePrim 234, Primitive4, "leftramp_Pink", "leftramp 0.28"
			nFadePrim 234, Primitive47, "heartramp_Blue_Lower", "heartramp 0.28b test"
			nFadePrim 235, Primitive47, "heartramp_Pink_Upper", "heartramp 0.28b test"
			If LampState(234) = 1 and LampState(235) = 1Then
				Primitive47.image = "heartramp_PinkBlue_AllOn"
			End If
			If LampState(204) = 0 Then
				nFadePrim 235, Primitive47, "heartramp 0.28b test", "heartramp 0.28b test"
				nFadePrim 235, Primitive4, "leftramp 0.28", "leftramp 0.28"
			End If

		End If


'					SetLamp 234, 1
'				Else
'					SetLamp 234, 0
'				End If
'			Case 4
'				If step >= 4 then
'					SetLamp 235, 1
'				Else
'					SetLamp 235, 0


    End Sub

'#end section
redim CollapseMe(0) 'Combined GI subs / functions (Click Me to Collapse)

    Set GICallback = GetRef("UpdateGIon")       'On/Off GI to NRs 200-203
    Sub UpdateGIOn(no, Enabled) : Setlamp no+200, cInt(enabled) : End Sub

    Set GICallback2 = GetRef("UpdateGI")
    Sub UpdateGI(no, step)                      '8 step Modulated GI to NRs 300-303
        Dim ii, x', i
        If step = 0 then exit sub 'only values from 1 to 8 are visible and reliable. 0 is not reliable and 7 & 8 are the same so...
        SetModLamp no+300, ScaleGI(step, 0)
        LampState((no+300)) = 0
    '   if no = 2 then tb.text = no & vbnewline & step & vbnewline & ScaleGI(step,0) & SolModValue(102)

		Select Case no
			Case 2
				If step >= 4 then
					SetLamp 234, 1
				Else
					SetLamp 234, 0

				End If
			Case 4
				If step >= 4 then
					SetLamp 235, 1
				Else
					SetLamp 235, 0

				End If
		End Select


    End Sub

    Function ScaleGI(value, scaletype)  'returns an intensityscale-friendly 0->100% value out of 1>8 'it does go to 8
        dim i
        Select Case scaletype   'select case because bad at maths
            case 0  : i = value * (1/8) '0 to 1
            case 25 : i = (1/28)*(3*value + 4)
            case 50 : i = (value+5)/12
            case else : i = value * (1/8)   '0 to 1
    '           x = (4*value)/3 - 85    '63.75 to 255
        End Select
        ScaleGI = i
    End Function

'   dim LSstate : LSstate = False   'fading sub handles SFX 'Uncomment to enable
    Sub FadeGI(nr) 'in On/off       'Updates nothing but flashlevel
        Select Case FadingLevel(nr)
            Case 3 : FadingLevel(nr) = 0
            Case 4 'off
    '           If Not LSstate then Playsound "FX_Relay_Off",0,LVL(0.1) : LSstate = True    'handle SFX
                FlashLevel(nr) = FlashLevel(nr) - (FlashSpeedDown(nr) * FrameTime)
                If FlashLevel(nr) < FlashMin(nr) Then
                   FlashLevel(nr) = FlashMin(nr)
                   FadingLevel(nr) = 3 'completely off
    '               LSstate = False
                End if
            Case 5 ' on
    '           If Not LSstate then Playsound "FX_Relay_On",0,LVL(0.1) : LSstate = True 'handle SFX
                FlashLevel(nr) = FlashLevel(nr) + (FlashSpeedUp(nr) * FrameTime)
                If FlashLevel(nr) > FlashMax(nr) Then
                    FlashLevel(nr) = FlashMax(nr)
                    FadingLevel(nr) = 6 'completely on
    '               LSstate = False
                End if
            Case 6 : FadingLevel(nr) = 1
        End Select
    End Sub
    Sub ModGI(nr2) 'in 0->1     'Updates nothing but flashlevel 'never off
        dim DesiredFading
        Select Case FadingLevel(nr2)
            case 3 : FadingLevel(nr2) = 0   'workaround - wait a frame to let M sub finish fading
    '       Case 4 : FadingLevel(nr2) = 3   'off -disabled off, only gicallback1 can turn off GI(?) 'experimental
            Case 5, 4 ' Fade (Dynamic)
                DesiredFading = SolModValue(nr2)
                if FlashLevel(nr2) < DesiredFading Then '+
                    FlashLevel(nr2) = FlashLevel(nr2) + (FlashSpeedUp(nr2)  * FrameTime )
                    If FlashLevel(nr2) >= DesiredFading Then FlashLevel(nr2) = DesiredFading : FadingLevel(nr2) = 1
                elseif FlashLevel(nr2) > DesiredFading Then '-
                    FlashLevel(nr2) = FlashLevel(nr2) - (FlashSpeedDown(nr2) * FrameTime    )
                    If FlashLevel(nr2) <= DesiredFading Then FlashLevel(nr2) = DesiredFading : FadingLevel(nr2) = 6
                End If
            Case 6
                FadingLevel(nr2) = 1
        End Select
    End Sub

    Sub UpdateGIobjects(nr, nr2, a) 'Just Update GI
        If FadingLevel(nr) > 1 or FadingLevel(nr2) > 1 Then
            dim x, Output : Output = FlashLevel(nr2) * FlashLevel(nr)
            for each x in a : x.IntensityScale = Output : next
        End If
    end Sub

    Sub GiCompensation(nr, nr2, a, GIscaleOff)  'One NR pairing only fading
    '   tbgi.text = "GI: " & SolModValue(nr) & " " & FlashLevel(nr) & " " & FadingLevel(nr) & vbnewline & _
    '               "ModGI: " & SolModValue(nr2) & " " & FlashLevel(nr2) & " " & FadingLevel(nr2) & vbnewline & _
    '               "Solmodvalue, Flashlevel, Fading step"
        if FadingLevel(nr) > 1 or FadingLevel(nr2) > 1 Then
            dim x, Giscaler, Output : Output = FlashLevel(nr2) * FlashLevel(nr)
            Giscaler = ((Giscaleoff-1) * (ABS(Output-1) )  ) + 1    'fade GIscale the opposite direction

            for x = 0 to (a.Count - 1) 'Handle Compensate Flashers
                On Error Resume Next
                a(x).Opacity = LampsOpacity(x, 0) * Giscaler
                a(x).Intensity = LampsOpacity(x, 0) * Giscaler
                a(x).FadeSpeedUp = LampsOpacity(x, 1) * Giscaler
                a(x).FadeSpeedDown = LampsOpacity(x, 2) * Giscaler
            Next
            '       tbbb.text = giscaler & " on:" & FadingLevel(nr) & vbnewline & "flash: " & output & " onmod:" & FadingLevel(nr2) & vbnewline & l37.intensity
            '       tbbb1.text = FadingLevel(nr) & vbnewline & FadingLevel(nr2)
            '   tbgi1.text = Output & " giscale:" & giscaler    'debug
        End If
        '       tbbb1.text = FLashLevel(nr) & vbnewline & FlashLevel(nr2)
    End Sub

    Sub GiCompensationAvg(nr, nr2, nr3, nr4, a, GIscaleOff) 'Two pairs of NRs averaged together
    '   tbgi.text = "GI: " & SolModValue(nr) & " " & FlashLevel(nr) & " " & FadingLevel(nr) & vbnewline & _
    '               "ModGI: " & SolModValue(nr2) & " " & FlashLevel(nr2) & " " & FadingLevel(nr2) & vbnewline & _
    '               "Solmodvalue, Flashlevel, Fading step"
        if FadingLevel(nr) > 1 or FadingLevel(nr2) > 1 or FadingLevel(nr3) > 1 or FadingLevel(nr4) > 1 Then
            dim x, Giscaler, Output : Output = (((FlashLevel(nr)*FlashLevel(nr2)) + (FlashLevel(nr3)*Flashlevel(nr4))) /2)
            Giscaler = ((Giscaleoff-1) * (ABS(Output-1) )  ) + 1    'fade GIscale the opposite direction

            for x = 0 to (a.Count - 1) 'Handle Compensate Flashers
                On Error Resume Next
                a(x).Opacity = LampsOpacity(x, 0) * Giscaler
                a(x).Intensity = LampsOpacity(x, 0) * Giscaler
                a(x).FadeSpeedUp = LampsOpacity(x, 1) * Giscaler
                a(x).FadeSpeedDown = LampsOpacity(x, 2) * Giscaler
            Next


        REM tbgi1.text = "Output:" & output & vbnewline & _
                    REM "GIscaler" & giscaler & vbnewline & _
                    REM "..."
        End If
        REM tbgi.text = "GI0 " & flashlevel(200) & " " & flashlevel(300) & vbnewline & _
                    REM "GI1 " & flashlevel(201) & " " & flashlevel(301) & vbnewline & _
                    REM "GI2 " & flashlevel(202) & " " & flashlevel(302) & vbnewline & _
                    REM "GI3 " & flashlevel(203) & " " & flashlevel(303) & vbnewline & _
                    REM "GI4 " & flashlevel(204) & " " & flashlevel(304) & vbnewline & _
                    REM "..."
    End Sub

    Sub GiCompensationAvgM(nr, nr2, nr3, nr4, nr5, nr6, a, GIscaleOff)  'Three pairs of NRs averaged together
    '   tbgi.text = "GI: " & SolModValue(nr) & " " & FlashLevel(nr) & " " & FadingLevel(nr) & vbnewline & _
    '               "ModGI: " & SolModValue(nr2) & " " & FlashLevel(nr2) & " " & FadingLevel(nr2) & vbnewline & _
    '               "Solmodvalue, Flashlevel, Fading step"
        if FadingLevel(nr) > 1 or FadingLevel(nr2) > 1 Then
            dim x, Giscaler, Output
            Output = (((FlashLevel(nr)*FlashLevel(nr2)) + (FlashLevel(nr3)*Flashlevel(nr4)) + (FlashLevel(nr5)*FlashLevel(nr6)))/3)

            Giscaler = ((Giscaleoff-1) * (ABS(Output-1) )  ) + 1    'fade GIscale the opposite direction

            for x = 0 to (a.Count - 1) 'Handle Compensate Flashers
                On Error Resume Next
                a(x).Opacity = LampsOpacity(x, 0) * Giscaler
                a(x).Intensity = LampsOpacity(x, 0) * Giscaler
                a(x).FadeSpeedUp = LampsOpacity(x, 1) * Giscaler
                a(x).FadeSpeedDown = LampsOpacity(x, 2) * Giscaler
            Next
            '       tbbb.text = giscaler & " on:" & FadingLevel(nr) & vbnewline & "flash: " & output & " onmod:" & FadingLevel(nr2) & vbnewline & l37.intensity
            '       tbbb1.text = FadingLevel(nr) & vbnewline & FadingLevel(nr2)
            '   tbgi1.text = Output & " giscale:" & giscaler    'debug
        End If
        '       tbbb1.text = FLashLevel(nr) & vbnewline & FlashLevel(nr2)
    End Sub

    Sub FadeLUT(nr, nr2, LutName, LutCount) 'fade lookuptable NOTE- this is a bad idea for darkening your table as
        If FadingLevel(nr) >2 or FadingLevel(nr2) > 2 Then              '-it will strip the whites out of your image
            dim GoLut
            GoLut = cInt(LutCount * (FlashLevel(nr)*FlashLevel(nr2) )   )
            Table.ColorGradeImage = LutName & GoLut
    '       tbgi2.text = Table.ColorGradeImage & vbnewline & golut  'debug
        End If
    End Sub

    Sub FadeLUTavg(nr, nr2, nr3, nr4, LutName, LutCount)    'FadeLut for two GI strings (WPC)
        If FadingLevel(nr) >2 or FadingLevel(nr2) > 2 or FadingLevel(nr3) >2 or FadingLevel(nr4) > 2 Then
            dim GoLut
            GoLut = cInt(LutCount * (((FlashLevel(nr)*FlashLevel(nr2)) + (FlashLevel(nr3)*Flashlevel(nr4))) /2) )
            Table.ColorGradeImage = LutName & GoLut
            REM tbgi2.text = Table.ColorGradeImage & vbnewline & golut  'debug
        End If
    End Sub

    Sub FadeLUTavgM(nr, nr2, nr3, nr4, nr5, nr6, LutName, LutCount) 'FadeLut for three GI strings (WPC)
        If FadingLevel(nr) >2 or FadingLevel(nr2) > 2 or FadingLevel(nr3) >2 or FadingLevel(nr4) > 2 or _
        FadingLevel(nr5) >2 or FadingLevel(nr6) > 2 Then
            dim GoLut
            GoLut = cInt(LutCount * (((FlashLevel(nr)*FlashLevel(nr2)) + (FlashLevel(nr3)*Flashlevel(nr4)) + (FlashLevel(nr5)*FlashLevel(nr6)))/3)  )   'what a mess
            Table.ColorGradeImage = LutName & GoLut
    '       tbgi2.text = Table.ColorGradeImage & vbnewline & golut  'debug
        End If
    End Sub

'#end section

redim CollapseMe(0) 'Fading subs     (Click Me to Collapse)
    Sub nModFlash(nr, object, scaletype, offscale)  'Fading with modulated callbacks
        dim DesiredFading
        Select Case FadingLevel(nr)
            case 3 : FadingLevel(nr) = 0    'workaround - wait a frame to let M sub finish fading
            Case 4  'off
                If Offscale = 0 then Offscale = 1
                FlashLevel(nr) = FlashLevel(nr) - (FlashSpeedDown(nr) * FrameTime   ) * offscale
                If FlashLevel(nr) < 0 then FlashLevel(nr) = 0 : FadingLevel(nr) = 3
                Object.IntensityScale = ScaleLights(FlashLevel(nr),0 )
            Case 5 ' Fade (Dynamic)
                DesiredFading = ScaleByte(SolModValue(nr), scaletype)
                if FlashLevel(nr) < DesiredFading Then '+
                    FlashLevel(nr) = FlashLevel(nr) + (FlashSpeedUp(nr) * FrameTime )
                    If FlashLevel(nr) >= DesiredFading Then FlashLevel(nr) = DesiredFading : FadingLevel(nr) = 1
                elseif FlashLevel(nr) > DesiredFading Then '-
                    FlashLevel(nr) = FlashLevel(nr) - (FlashSpeedDown(nr) * FrameTime   )
                    If FlashLevel(nr) <= DesiredFading Then FlashLevel(nr) = DesiredFading : FadingLevel(nr) = 6
                End If
                Object.Intensityscale = ScaleLights(FlashLevel(nr),0 )
            Case 6 : FadingLevel(nr) = 1
        End Select
    End Sub

    Sub nModFlashM(nr, Object)
        Select Case FadingLevel(nr)
            Case 3, 4, 5, 6 : Object.Intensityscale = ScaleLights(FlashLevel(nr),0 )
        End Select
    End Sub

    Sub Flashc(nr, object)  'FrameTime Compensated. Can work with Light Objects (make sure state is 1 though)
        Select Case FadingLevel(nr)
            Case 3 : FadingLevel(nr) = 0
            Case 4 'off
                FlashLevel(nr) = FlashLevel(nr) - (FlashSpeedDown(nr) * FrameTime)
                If FlashLevel(nr) < FlashMin(nr) Then
                    FlashLevel(nr) = FlashMin(nr)
                   FadingLevel(nr) = 3 'completely off
                End if
                Object.IntensityScale = FlashLevel(nr)
            Case 5 ' on
                FlashLevel(nr) = FlashLevel(nr) + (FlashSpeedUp(nr) * FrameTime)
                If FlashLevel(nr) > FlashMax(nr) Then
                    FlashLevel(nr) = FlashMax(nr)
                    FadingLevel(nr) = 6 'completely on
                End if
                Object.IntensityScale = FlashLevel(nr)
            Case 6 : FadingLevel(nr) = 1
        End Select
    End Sub

    Sub Flashm(nr, object) 'multiple flashers, it just sets the flashlevel
        select case FadingLevel(nr)
            case 3, 4, 5, 6 : Object.IntensityScale = FlashLevel(nr)
        end select
    End Sub

    Sub NFadeL(nr, object)  'Simple VPX light fading using State
   Select Case FadingLevel(nr)
        Case 3:object.state = 0:FadingLevel(nr) = 0
        Case 4:object.state = 0:FadingLevel(nr) = 3
        Case 5:object.state = 1:FadingLevel(nr) = 6
        Case 6:object.state = 1:FadingLevel(nr) = 1
    End Select
    End Sub

    Sub NFadeLm(nr, object) ' used for multiple lights
        Select Case FadingLevel(nr)
            Case 3:object.state = 0
            Case 4:object.state = 0
            Case 5:object.state = 1
            Case 6:object.state = 1
        End Select
    End Sub

    Sub NFadeLFO(nr)    'Old stuff
        Select Case FadingLevel(nr)
            Case 3:FadingLevel(nr) = 0
            Case 4:SetLamp nr, 0:FadingLevel(nr) = 3
            Case 5:SetLamp nr, 1:FadingLevel(nr) = 6
            Case 6:FadingLevel(nr) = 1
        End Select
    End Sub

    Sub NFadeLFOm(nr)
        Select Case FadingLevel(nr)
            Case 3,4:SetLamp nr, 0
            Case 5,6:SetLamp nr, 1
        End Select
    End Sub


    Sub NFadePri(nr, pri, a, b)
    Select Case FadingLevel(nr)
        Case 3:pri.image = b:FadingLevel(nr) = 0 'off
       Case 4:pri.image = b:FadingLevel(nr) = 3 'off
       Case 5:pri.image = a:FadingLevel(nr) = 6 'on
        Case 6:pri.image = a:FadingLevel(nr) = 1 'on
   End Select
    End Sub

    Sub NFadePrim(nr, pri, a, b)
        Select Case FadingLevel(nr)
            Case 3, 4:pri.image = b
            Case 5, 6:pri.image = a
        End Select
    End Sub

    'Toggles disable lighting on and off -cp
    Sub FadeDisableLighting(nr, a)
        Select Case FadingLevel(nr)
            Case 3, 4:a.DisableLighting = 0
            Case 5, 6:a.DisableLighting = 1
        End Select
    End Sub

'#End Section

redim CollapseMe(0) 'Fading Functions (Click Me to Collapse)
    Function ScaleLights(value, scaletype)  'returns an intensityscale-friendly 0->100% value out of 255
        dim i
        Select Case scaletype   'select case because bad at maths   'TODO: Simplify these functions. B/c this is absurdly bad.
            case 0  : i = value * (1 / 255) '0 to 1
            case 6  : i = (value + 17)/272  '0.0625 to 1
            case 9  : i = (value + 25)/280  '0.089 to 1
            case 15 : i = (value / 300) + 0.15
            case 20 : i = (4 * value)/1275 + (1/5)
            case 25 : i = (value + 85) / 340
            case 37 : i = (value+153) / 408     '0.375 to 1
            case 40 : i = (value + 170) / 425
            case 50 : i = (value + 255) / 510   '0.5 to 1
            case 75 : i = (value + 765) / 1020  '0.75 to 1
            case Else : i = 10
        End Select
        ScaleLights = i
    End Function

    Function ScaleByte(value, scaletype)    'returns a number between 1 and 255
        dim i
        Select Case scaletype
            case 0 : i = value * 1  '0 to 1
            case 9 : i = (5*(200*value + 1887))/1037 'ugh
            case 15 : i = (16*value)/17 + 15
            Case 63 : i = (3*(value + 85))/4
            case else : i = value * 1   '0 to 1
        End Select
        ScaleByte = i
    End Function

'#end section

'.... End Light Routine........
'==============================

Sub UpdateLampsOld

End Sub

'*************************************
'          Nudge System
' JP's based on Noah's nudgetest table
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
    If LeftNudgeEffect = 0 then LeftNudgeTimer.Enabled = 0
End Sub

Sub RightNudgeTimer_Timer()
    RightNudgeEffect = RightNudgeEffect-1
    If RightNudgeEffect = 0 then RightNudgeTimer.Enabled = 0
End Sub

Sub NudgeTimer_Timer()
    NudgeEffect = NudgeEffect-1
    If NudgeEffect = 0 then NudgeTimer.Enabled = 0
End Sub

' ******************************************************************************************************************************************
'**********Sling Shot Animations
'****************
Dim Ustep

Sub UpperSlingShot_Slingshot
    PlaySound SoundFX("fx_slingshot_r",DOFContactors),0,1,-0.05,0.05
    USling.Visible = 0
    USling1.Visible = 1
    slingu.TransZ = -20
    UStep = 0
    UpperSlingShot.TimerEnabled = 1
    vpmTimer.PulseSw 56
    'gi1.State = 0:Gi2.State = 0
End Sub

Sub UpperSlingShot_Timer
    Select Case UStep
        Case 1:USLing1.Visible = 0:USLing2.Visible = 1:slingu.transZ = -10
        Case 2:USLing2.Visible = 0:USLing.Visible = 1:slingu.transZ = 0:UpperSlingShot.TimerEnabled = 0:'gi1.State = 1:Gi2.State = 1
   End Select
    UStep = UStep + 1
End Sub

Dim LStep, RStep

Sub LeftSlingShot_Slingshot
    PlaySound SoundFX("fx_slingshot_l",DOFContactors),0,1,-0.05,0.05:vpmTimer.PulseSw 57
    LSling.Visible = 0
    LSling3.Visible = 1
    sling2.TransZ = -20
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
    vpmTimer.PulseSw 57
    'gi1.State = 0:Gi2.State = 0
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 1:LSLing3.Visible = 0:LSLing4.Visible = 1:sling2.TransZ = -10
        Case 2:LSLing4.Visible = 0:LSLing.Visible = 1:sling2.TransZ = 0:LeftSlingShot.TimerEnabled = 0:'gi1.State = 1:Gi2.State = 1
   End Select
    LStep = LStep + 1
End Sub

Sub RightSlingShot_Slingshot
    PlaySound SoundFX("fx_slingshot_r",DOFContactors), 0, 1, 0.05, 0.05:vpmTimer.PulseSw 58
    RSling.Visible = 0
    RSling3.Visible = 1
    sling1.TransZ = -20
    RStep = 0
    RightSlingShot.TimerEnabled = 1
    vpmTimer.PulseSw 58
    'gi1.State = 0:Gi2.State = 0
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 1:RSLing3.Visible = 0:RSLing4.Visible = 1:sling1.TransZ = -10
        Case 2:RSLing4.Visible = 0:RSLing.Visible = 1:sling1.TransZ = 0:RightSlingShot.TimerEnabled = 0:'gi1.State = 1:Gi2.State = 1
   End Select
    RStep = RStep + 1
End Sub

'**********************************************************************************************************
'Digital Display
'**********************************************************************************************************
Dim Digits(32)
 Digits(0)=Array(a00, a05, a0c, a0d, a08, a01, a06, a0f, a02, a03, a04, a07, a0b, a0a, a09, a0e)
 Digits(1)=Array(a10, a15, a1c, a1d, a18, a11, a16, a1f, a12, a13, a14, a17, a1b, a1a, a19, a1e)
 Digits(2)=Array(a20, a25, a2c, a2d, a28, a21, a26, a2f, a22, a23, a24, a27, a2b, a2a, a29, a2e)
 Digits(3)=Array(a30, a35, a3c, a3d, a38, a31, a36, a3f, a32, a33, a34, a37, a3b, a3a, a39, a3e)
 Digits(4)=Array(a40, a45, a4c, a4d, a48, a41, a46, a4f, a42, a43, a44, a47, a4b, a4a, a49, a4e)
 Digits(5)=Array(a50, a55, a5c, a5d, a58, a51, a56, a5f, a52, a53, a54, a57, a5b, a5a, a59, a5e)
 Digits(6)=Array(a60, a65, a6c, a6d, a68, a61, a66, a6f, a62, a63, a64, a67, a6b, a6a, a69, a6e)
 Digits(7)=Array(a70, a75, a7c, a7d, a78, a71, a76, a7f, a72, a73, a74, a77, a7b, a7a, a79, a7e)
 Digits(8)=Array(a80, a85, a8c, a8d, a88, a81, a86, a8f, a82, a83, a84, a87, a8b, a8a, a89, a8e)
 Digits(9)=Array(a90, a95, a9c, a9d, a98, a91, a96, a9f, a92, a93, a94, a97, a9b, a9a, a99, a9e)
 Digits(10)=Array(aa0, aa5, aac, aad, aa8, aa1, aa6, aaf, aa2, aa3, aa4, aa7, aab, aaa, aa9, aae)
 Digits(11)=Array(ab0, ab5, abc, abd, ab8, ab1, ab6, abf, ab2, ab3, ab4, ab7, abb, aba, ab9, abe)
 Digits(12)=Array(ac0, ac5, acc, acd, ac8, ac1, ac6, acf, ac2, ac3, ac4, ac7, acb, aca, ac9, ace)
 Digits(13)=Array(ad0, ad5, adc, add, ad8, ad1, ad6, adf, ad2, ad3, ad4, ad7, adb, ada, ad9, ade)
 Digits(14)=Array(ae0, ae5, aec, aed, ae8, ae1, ae6, aef, ae2, ae3, ae4, ae7, aeb, aea, ae9, aee)
 Digits(15)=Array(af0, af5, afc, afd, af8, af1, af6, aff, af2, af3, af4, af7, afb, afa, af9, afe)
 Digits(16)=Array(b00, b05, b0c, b0d, b08, b01, b06, b0f, b02, b03, b04, b07, b0b, b0a, b09, b0e)
 Digits(17)=Array(b10, b15, b1c, b1d, b18, b11, b16, b1f, b12, b13, b14, b17, b1b, b1a, b19, b1e)
 Digits(18)=Array(b20, b25, b2c, b2d, b28, b21, b26, b2f, b22, b23, b24, b27, b2b, b2a, b29, b2e)
 Digits(19)=Array(b30, b35, b3c, b3d, b38, b31, b36, b3f, b32, b33, b34, b37, b3b, b3a, b39, b3e)
 Digits(20)=Array(b40, b45, b4c, b4d, b48, b41, b46, b4f, b42, b43, b44, b47, b4b, b4a, b49, b4e)
 Digits(21)=Array(b50, b55, b5c, b5d, b58, b51, b56, b5f, b52, b53, b54, b57, b5b, b5a, b59, b5e)
 Digits(22)=Array(b60, b65, b6c, b6d, b68, b61, b66, b6f, b62, b63, b64, b67, b6b, b6a, b69, b6e)
 Digits(23)=Array(b70, b75, b7c, b7d, b78, b71, b76, b7f, b72, b73, b74, b77, b7b, b7a, b79, b7e)
 Digits(24)=Array(b80, b85, b8c, b8d, b88, b81, b86, b8f, b82, b83, b84, b87, b8b, b8a, b89, b8e)
 Digits(25)=Array(b90, b95, b9c, b9d, b98, b91, b96, b9f, b92, b93, b94, b97, b9b, b9a, b99, b9e)
 Digits(26)=Array(ba0, ba5, bac, bad, ba8, ba1, ba6, baf, ba2, ba3, ba4, ba7, bab, baa, ba9, bae)
 Digits(27)=Array(bb0, bb5, bbc, bbd, bb8, bb1, bb6, bbf, bb2, bb3, bb4, bb7, bbb, bba, bb9, bbe)
 Digits(28)=Array(bc0, bc5, bcc, bcd, bc8, bc1, bc6, bcf, bc2, bc3, bc4, bc7, bcb, bca, bc9, bce)
 Digits(29)=Array(bd0, bd5, bdc, bdd, bd8, bd1, bd6, bdf, bd2, bd3, bd4, bd7, bdb, bda, bd9, bde)
 Digits(30)=Array(be0, be5, bec, bed, be8, be1, be6, bef, be2, be3, be4, be7, beb, bea, be9, bee)
 Digits(31)=Array(bf0, bf5, bfc, bfd, bf8, bf1, bf6, bff, bf2, bf3, bf4, bf7, bfb, bfa, bf9, bfe)

 Sub DisplayTimer_Timer
    Dim ChgLED, ii, jj, num, chg, stat, obj, b, x
    ChgLED=Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
    If Not IsEmpty(ChgLED)Then
        If DesktopMode = True Then
       For ii=0 To UBound(chgLED)
          num=chgLED(ii, 0) : chg=chgLED(ii, 1) : stat=chgLED(ii, 2)
            if (num < 32) then
              For Each obj In Digits(num)
                   If chg And 1 Then obj.State=stat And 1
                   chg=chg\2 : stat=stat\2
                  Next
            Else
                   end if
        Next
       end if
    End If
 End Sub
'**********************************************************************************************************
'**********************************************************************************************************

Sub RWireStart_Hit():PlaySound "fx_metalrolling":End Sub

Sub Gates_Hit (idx)
    PlaySound "gate", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub BallHitSound(dummy):PlaySound "ball_bounce":End Sub


Sub RWireStart_Hit()
If ActiveBall.VelY < 0 Then Playsound "fx_metalrolling"
End Sub

Sub plungeballdrop_Hit()
If ActiveBall.VelY > 0 Then Playsound "wirerampdrop"
End Sub

Sub RWireEnd_Hit()
     vpmTimer.AddTimer 150, "BallHitSound"
     StopSound "fx_metalrolling"
 End Sub

Sub Trigger1_Hit():PlaySound "Ramp":End Sub
Sub Trigger2_Hit():PlaySound "Ramp":End Sub
Sub Trigger3_Hit():PlaySound "Ramp":End Sub
Sub Trigger4_Hit():PlaySound "Ramp":End Sub
Sub Trigger5_Hit():PlaySound "Ramp":End Sub
Sub Trigger6_Hit():PlaySound "Ramp" End Sub
Sub Trigger7_Hit():PlaySound "Ramp":End Sub
Sub Trigger8_Hit():PlaySound "Ramp":End Sub
Sub Trigger9_Hit():PlaySound "Ramp":End Sub
Sub Trigger10_Hit():PlaySound "Ramp":End Sub

Sub RubberslowerPFlargepost_Hit(idx)
    dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
    If finalspeed > 3 then
        PlaySound "fx_rubber2", 0, 2*Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    End if
    If finalspeed >= 6 AND finalspeed <= 20 then
        RandomSoundRubber()
    End If
End Sub

Sub RubberlowerPFsmallpost_Hit(idx)
    dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
    If finalspeed > 3 then
        PlaySound "rubber", 0, 2*Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
    End if
    If finalspeed >= 6 AND finalspeed <= 16 then
        RandomSoundRubber()
    End If
End Sub





Sub ShuttleRampStart_Hit:PlaySound "plasticrollinglowhighpass", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0: End Sub
Sub ShuttleRampEnd_UnHit(): StopSound "plasticrollinglowhighpass" :End Sub
Sub ShuttleRampEnd1_Hit(): StopSound "plasticrollinglowhighpass" : End Sub
Sub ShuttleRampEnd2_Hit(): StopSound "plasticrollinglowhighpass" : End Sub

Sub HeartRampStart_Hit ():PlaySound "plasticrollinglowhighpass", 0, 50*Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0: End Sub
Sub HeartRampEnd_UnHit(): StopSound "plasticrollinglowhighpass" :End Sub

Sub HeartRampStart1_Hit:PlaySound "plasticrollinglowhighpass", 0, 50*Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0: End Sub
Sub HeartRampEnd1_UnHit(): StopSound "plasticrollinglowhighpass" :End Sub


'MISC. Sound effects

Sub Subway_Hit(): playsound"fx_metalrolling" : End Sub
Sub UPFdrop_Hit: playsound"ball_bounce" : End Sub
Sub xrampdrop_Hit: playsound"ball_bounce" : End Sub
Sub lrampdrop_Hit: playsound"ball_bounce" : End Sub
Sub BL_hit():PlaySound "wirerampdrop":End Sub
Sub right_gate_Hit():if gate3.currentangle < 90 then:PlaySound "fx_gate2":end if:End Sub
Sub Metals_Hit(idx):PlaySound "fx_metalhit2", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub Gates_Hit (idx): PlaySound "fx_gate", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall): End Sub

Sub ramphelper_Hit()
	ActiveBall.vely = Activeball.vely*2.5
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''  Ball Trough system ''''''''''''''''''''''''''
'''''''''''''''''''''by cyberpez''''''''''''''''''''''''''''''''
''''''''''''''''based off of EalaDubhSidhe's''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim BallCount

Sub CheckMaxBalls()
    BallCount = MaxBalls
    TroughWall1.isDropped = true
    TroughWall2.isDropped = true
End Sub

Sub CreatBalls_timer()
    If BallCount > 0 then
        drain.CreateSizedBallWithMass BallSize, BallMass
        Drain.kick 70,20
        BallCount = BallCount - 1
    End If

    If BallCount = 0 Then
        CreatBalls.enabled = false

    End If
End Sub

Dim DRSstep

Sub DelayRollingStart_timer()
    Select Case DRSstep
        Case 5: RollingSoundTimer.enabled = true
    End Select
    DRSstep = DRSstep + 1
End Sub

Sub ballrelease_hit()
'   Kicker1active = 1
    Controller.Switch(25)=1
    TroughWall1.isDropped = false

End Sub

Sub sw26_Hit()
    Controller.Switch(26)=1
    TroughWall2.isDropped = false
End Sub

Sub sw26_unHit()
    Controller.Switch(26)=0
    TroughWall2.isDropped = true
End Sub

Sub sw27_Hit()
    Controller.Switch(27)=1
End Sub

Sub sw27_unHit()
    Controller.Switch(27)=0
End Sub

Sub KickBallToLane(Enabled)
    PlaySound SoundFX("BallRelease",DOFContactors)
    PlaySound SoundFX("Solenoid",DOFContactors)
    ballrelease.Kick 60,10
    TroughWall1.isDropped = true
    Controller.Switch(25)=0
End Sub


sub kisort(enabled)
    Drain.Kick 70,20
    controller.switch(38) = false
end sub


Sub Drain_hit()
    PlaySound "drain"
    controller.switch(38) = true
End Sub



' _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _
'(_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_)


'__          __   _____       _   _                   _                     __          __
'\ \        / /  |  _  |     | | (_)                 | |                    \ \        / /
' \ \      / /   | | | |_ __ | |_ _  ___  _ __  ___  | |__   ___ _ __ ___    \ \      / /
'  \ \    / /    | | | | '_ \| __| |/ _ \| '_ \/ __| | '_ \ / _ \ '__/ _ \    \ \    / /
'   \ \  / /     \ \_/ / |_) | |_| | (_) | | | \__ \ | | | |  __/ | |  __/     \ \  / /
'    \_\/_/       \___/| .__/ \__|_|\___/|_| |_|___/ |_| |_|\___|_|  \___|      \_\/_/
'                      | |
'                      |_|

Dim TableOptions, TableName, Rubbercolor,xxrubbercolor, RubberColorType, GIColorMod, GIColorModType, SSRampMod, SSRampModType, HLColorType, HottieMod, HottieModType, ChooseBats, ChooseBatsType

Private vpmShowDips1, vpmDips1

Sub InitializeOptions
    TableName="BOP_VPX"                                         'Your descriptive table name, it will be used to save settings in VPReg.stg file
    Set vpmShowDips1 = vpmShowDips                              'Reassigns vpmShowDips to vpmShowDips1 to allow usage of default dips menu
    Set vpmShowDips = GetRef("TableShowDips")                   'Assigns new sub to vmpShowDips
    TableOptions = LoadValue(TableName,"Options")               'Load saved table options

    Set Controller = CreateObject("VPinMAME.Controller")        'Load vpm controller temporarily so options menu can be loaded if needed
    If TableOptions = "" Or optReset Then                       'If no existing options, reset to default through optReset, then open Options menu
        TableOptions = 1                                        'clear any existing settings and set table options to default options
        TableShowOptions
    ElseIf (TableOptions And 1) Then                            'If Enable Next Start was selected then
        'TableOptions = TableOptions - 1                        'clear setting to avoid future executions
        TableShowOptions
    Else
        TableSetOptions
    End If

    Set Controller = Nothing                                    'Unload vpm controller so selected controller can be loaded

End Sub

Private Sub TableShowDips
    vpmShowDips1                                                'Show original Dips menu
    TableShowOptions                                            'Show new options menu
End Sub

Private Sub TableShowOptions                    'New options menu
    Dim oldOptions : oldOptions = TableOptions
    If Not IsObject(vpmDips1) Then
        Set vpmDips1 = New cvpmDips
        With vpmDips1
            .AddForm 700, 500, "TABLE OPTIONS MENU"
            .AddFrameExtra 0,0,155,"Space Shuttle Toy Mod",2^1, Array("Not Visible", 0, "Visible", 2^1)
            .AddFrameExtra 0,48,155,"Sidewalls Type Mod",2^2+2^3+2^4, Array("Random", 0, "Wood Walls", 2^2, "Starfield Walls", 2^3, "BOP CabArt Style", 2^4)
            .AddFrameExtra 0,123,155,"Cheater Post Mod",2^5, Array("No Drain Post", 0, "Add Drain Post", 2^5)
            .AddFrameExtra 0,171,155,"Helmet Lights Color Mod",2^6+2^17, Array("Random lights", 0, "Normal lights", 2^6, "2.0 (blue) lights", 2^17)
            .AddFrameExtra 0,232,155,"Music Intro Mod",2^7, Array("Intro On", 0, "Intro Off", 2^7)
            .AddFrameExtra 0,279,155,"Rubbers Color Mod",2^8+2^9, Array("Random", 0, "Black", 2^8, "White", 2^9)
			.AddFrameExtra 0,341,155,"Helmet Lights Reflections",2^10, Array("None", 0, "Ball Reflections", 2^10)
			.AddFrameExtra 175,0,155,"GI Color Mod",2^11+2^12+2^13+2^14, Array("Random", 0, "Normal", 2^11, "All Blue", 2^12, "All Pink", 2^13, "Blue/Pink", 2^14)
            .AddFrameExtra 175,90,155,"SkillShot Ramp Mod",2^15+2^16, Array("Random", 0, "Normal", 2^15, "Color", 2^16)
            .AddFrameExtra 175,152,155,"Hottie Mod",2^17+2^18+2^19+2^20+2^21+2^22+2^23+2^24+2^25+2^26, Array("Random", 0, "Normal BOP face", 2^17, "Ashley", 2^18, "Brooke", 2^19, "Harley", 2^26, "Kelly", 2^20, "Lisa", 2^21, "Meagan", 2^22, "Shelly", 2^23, "Rachel", 2^24, "Valerie", 2^25)
            .AddFrameExtra 175,328,155,"Flipper Rubbers Color Mod",2^27+2^28+2^29+2^30, Array("Random", 0, "Black", 2^28, "Red", 2^29, "Blue", 2^30)
              .AddChkExtra 0,412,155, Array("Hide Menu Next Start", 1)
			     .Addlabel 0,440,155,21,"To re-enable hidden Startup Options Menu"
			     .Addlabel 0,455,155,21,"Refer to Table Script."
        End With
End If

    TableOptions = vpmDips1.ViewDipsExtra(TableOptions)
    SaveValue TableName,"Options",TableOptions
    TableSetOptions
    SetOptions
End Sub

Sub TableSetOptions     'defines required settings before table is run
    ShipMod = (TableOptions And 2^1):If ShipMod = 2^1 Then Shipmod = 1
    sidewalls = (TableOptions And 2^2+2^3+2^4)
        If sidewalls = 2^2 then
            sidewalls = 1
        Elseif sidewalls = 2^3 then
            sidewalls = 2
        Elseif sidewalls = 2^4 then
            sidewalls = 3
        End If

    cheaterpost = (TableOptions And 2^5): If cheaterpost = 2^5 Then cheaterpost = 1


    HLcolor = (TableOptions And 2^6+2^17)
		If HLcolor = 2^6 Then
			HLcolor = 1
		Elseif HLcolor = 2^17 Then
			HLcolor = 2
		End If

    RubberColor = (TableOptions And 2^8+2^9)
		If RubberColor = 2^8 Then
			RubberColor = 1
		Elseif RubberColor = 2^9 Then
			RubberColor = 2
		End If

    Musicsnippet = (TableOptions And 2^7): If Musicsnippet = 2^7 Then Musicsnippet = 1

	helmetreflections = (TableOptions And 2^10):If helmetreflections = 2^10 then helmetreflections = 1

    GIColorMod = (TableOptions And 2^11+2^12+2^13+2^14)
		If GIColorMod = 2^11 Then
			GIColorMod = 1
		Elseif GIColorMod = 2^12 Then
			GIColorMod = 2
		Elseif GIColorMod = 2^13 Then
			GIColorMod = 3
		Elseif GIColorMod = 2^14 Then
			GIColorMod = 4
		End If

    SSRampMod = (TableOptions And 2^15+2^16)
		If SSRampMod = 2^15 Then
			SSRampMod = 1
		Elseif SSRampMod = 2^16 Then
			SSRampMod = 2
		End If

    HottieMod = (TableOptions And 2^17+2^18+2^19+2^20+2^21+2^22+2^23+2^24+2^25+2^26)
		If HottieMod = 2^17 Then
			HottieMod = 1
		Elseif HottieMod = 2^18 Then
			HottieMod = 2
		Elseif HottieMod = 2^19 Then
			HottieMod = 3
		Elseif HottieMod = 2^20 Then
			HottieMod = 4
		Elseif HottieMod = 2^21 Then
			HottieMod = 5
		Elseif HottieMod = 2^22 Then
			HottieMod = 6
		Elseif HottieMod = 2^23 Then
			HottieMod = 7
		Elseif HottieMod = 2^24 Then
			HottieMod = 8
		Elseif HottieMod = 2^25 Then
			HottieMod = 9
		Elseif HottieMod = 2^26 Then
			HottieMod = 10
		End If

    ChooseBats = (TableOptions And 2^28+2^29+2^30)
		If ChooseBats = 2^28 Then
			ChooseBats = 1
		Elseif ChooseBats = 2^29 Then
			ChooseBats = 2
		Elseif ChooseBats = 2^30 Then
			ChooseBats = 3
		End If


    SaveValue TableName,"Options",TableOptions
    SetOptions
End Sub

Dim DivPOS, SideWallType, HLcolor

'''''''' SetOptions
Sub SetOptions()

'ShipMod
    If ShipMod = 1 Then
        pShipToy.Visible = True
        F124c.visible = true
        F124d.visible = true
        F124e.visible = true
        F124f.visible = true
        F124g.visible = true
        F124h.visible = true
    Else
        pShipToy.Visible = False
        F124c.visible = False
        F124d.visible = False
        F124e.visible = False
        F124f.visible = False
        F124g.visible = False
        F124h.visible = False
    End if

'Cheaters
    If cheaterpost = 1 then
        crubber.IsDropped = 0
        crubber.visible = 1
        cpost.visible = 1
    Else
        cpost.Z = -120
        crubber.IsDropped = 1
        crubber.visible = 0
    End If

' Sidewall switching
    If sidewalls = 0 then
        SideWallType = Int(Rnd*3)+1
    Else
        SideWallType = sidewalls
    End If


    select case SideWallType
        case 1: pSidewall.image = "sidewalls_texture3":pSidewall.visible = true
        case 2: pSidewall.image = "sidewalls_texture2":pSidewall.visible = true
        case 3: pSidewall.image = "sidewalls_texture":pSidewall.visible = true
    end select

''''Random Diverter POS at startup
    DivPOS = Int(Rnd*2)+1
    If DivPOS = 1 Then
        div.RotateToStart
    Else
        div.RotateToEnd
    End If

'Helmet Lamp Color
Dim BlueFull, Blue, BlueI, BlueFlasher, YellowFull, Yellow, YellowI, YellowFlasher, xxHLColorL, xxHLColorF, xxHLRefl

BlueFull = rgb(0,255,255)
Blue = rgb(150,150,255)
BlueFlasher = rgb(0,255,255)
BlueI = 20

YellowFull = rgb(255,255,255)
Yellow = rgb(255,255,128)
YellowFlasher = rgb(255,255,230)
YellowI = 20

If HLColor = 0 Then
	HLColorType = Int(Rnd*2)+1
Else
	HLColorType = HLColor
End If

If HLColorType = 1 Then
for each xxHLColorL in HelmetLights_FS
xxHLColorL.Color=Yellow
xxHLColorL.ColorFull=YellowFull
xxHLColorL.Intensity = yellowi
next
for each xxHLColorF in HelmetFlashers_FS
xxHLColorF.Color=YellowFlasher
next
End If

If HLColorType = 2 Then
for each xxHLColorL in HelmetLights_FS
xxHLColorL.Color=Blue
xxHLColorL.ColorFull=BlueFull
xxHLColorL.Intensity = BlueI
next
for each xxHLColorF in HelmetFlashers_FS
xxHLColorF.Color=BlueFlasher
next
End If

If RubberColor = 0 Then
	RubberColorType = Int(Rnd*2)+1
Else
	RubberColorType = RubberColor
End If


If RubberColorType = 1 Then
	for each xxRubberColor in LowerRubbersBandsLargeRings
		xxRubberColor.Material = "RubberBlack"
		next
	for each xxRubberColor in RubbersSmallRings
		xxRubberColor.Material = "RubberBlack"
		next
	for each xxRubberColor in UpperRubbersSmallRings
		xxRubberColor.Material = "RubberBlack"
		next
	for each xxRubberColor in UpperRubbersBandsLargeRings
		xxRubberColor.Material = "RubberBlack"
		next
End If


If RubberColorType = 2 Then
	for each xxRubberColor in LowerRubbersBandsLargeRings
		xxRubberColor.Material = "RubberWhite"
		next
	for each xxRubberColor in RubbersSmallRings
		xxRubberColor.Material = "RubberWhite"
		next
	for each xxRubberColor in UpperRubbersSmallRings
		xxRubberColor.Material = "RubberWhite"
		next
	for each xxRubberColor in UpperRubbersBandsLargeRings
		xxRubberColor.Material = "RubberWhite"
		next
End If

If helmetreflections = 1 Then
	for each xxHLRefl in HelmetLights_FS
		xxHLRefl.ShowReflectionOnBall = True
	next
Else
	for each xxHLRefl in HelmetLights_FS
		xxHLRefl.ShowReflectionOnBall = False
	next
End If

Dim xxGIColor, Red, RedFull, RedI, Pink, PinkFull, PinkI, White, WhiteFull, WhiteI, Blue2, Blue2Full, Blue2I, Yellow2, Yellow2Full, Yellow2I, Orange, OrangeFull, OrangeI

RedFull = rgb(255,255,255)
Red = rgb(255,0,0)
RedI = 5
PinkFull = rgb(255,0,225)
Pink = rgb(255,0,255)
PinkI = 10
WhiteFull = rgb(255,255,128)
White = rgb(255,255,255)
WhiteI = 7
Blue2Full = rgb(0,100,255)
Blue2 = rgb(0,255,255)
Blue2I = 20
Yellow2Full = rgb(255,255,128)
Yellow2 = rgb(255,255,0)
Yellow2I = 20
OrangeFull = rgb(255,128,64)
Orange = rgb(128,128,0)
OrangeI = 20

If GIColorMod = 0 Then
	GIColorModType = Int(Rnd*4)+1
Else
	GIColorModType = GIColorMod
End If

If GIColorModType = 2 then
	for each xxGIColor in GIColorA
		xxGIColor.Color=Blue2
		xxGIColor.ColorFull=Blue2Full
		xxGIColor.Intensity = Blue2I
		next
	for each xxGIColor in GIColorB
		xxGIColor.Color=Blue2
		xxGIColor.ColorFull=Blue2Full
		xxGIColor.Intensity = Blue2I
		next
End If
If GIColorModType = 3 then
	for each xxGIColor in GIColorA
		xxGIColor.Color=Pink
		xxGIColor.ColorFull=PinkFull
		xxGIColor.Intensity = PinkI
		next
	for each xxGIColor in GIColorB
		xxGIColor.Color=Pink
		xxGIColor.ColorFull=PinkFull
		xxGIColor.Intensity = PinkI
		next
End If
If GIColorModType = 4 then
	for each xxGIColor in GIColorA
		xxGIColor.Color=Blue2
		xxGIColor.ColorFull=Blue2Full
		xxGIColor.Intensity = Blue2I
		next
	for each xxGIColor in GIColorB
		xxGIColor.Color=Pink
		xxGIColor.ColorFull=PinkFull
		xxGIColor.Intensity = PinkI
		next
End If

If SSRampMod = 0 Then
	SSRampModType = Int(Rnd*2)+1
Else
	SSRampModType = SSRampMod
End If

If SSRampModType = 1 Then
		Light23b.Color=White
		Light23b.ColorFull=WhiteFull
		Light23b.Intensity=WhiteI
		Light23b1.Color=White
		Light23b1.ColorFull=WhiteFull
		Light23b1.Intensity=WhiteI
		Light23b2.Color=White
		Light23b2.ColorFull=WhiteFull
		Light23b2.Intensity=WhiteI

		Light22b.Color=White
		Light22b.ColorFull=WhiteFull
		Light22b.Intensity=WhiteI
		Light22b1.Color=White
		Light22b1.ColorFull=WhiteFull
		Light22b1.Intensity=WhiteI
		Light22b2.Color=White
		Light22b2.ColorFull=WhiteFull
		Light22b2.Intensity=WhiteI

		Light21b.Color=White
		Light21b.ColorFull=WhiteFull
		Light21b.Intensity=WhiteI
		Light21b1.Color=White
		Light21b1.ColorFull=WhiteFull
		Light21b1.Intensity=WhiteI
		Light21b2.Color=White
		Light21b2.ColorFull=WhiteFull
		Light21b2.Intensity=WhiteI

		Light20b.Color=White
		Light20b.ColorFull=WhiteFull
		Light20b.Intensity=WhiteI
		Light20b1.Color=White
		Light20b1.ColorFull=WhiteFull
		Light20b1.Intensity=WhiteI
		Light20b2.Color=White
		Light20b2.ColorFull=WhiteFull
		Light20b2.Intensity=WhiteI

		Light19b.Color=White
		Light19b.ColorFull=WhiteFull
		Light19b.Intensity=WhiteI
		Light19b1.Color=White
		Light19b1.ColorFull=WhiteFull
		Light19b1.Intensity=WhiteI
		Light19b2.Color=White
		Light19b2.ColorFull=WhiteFull
		Light19b2.Intensity=WhiteI
End If


If SSRampModType = 2 Then
		Light23b.Color=Blue2
		Light23b.ColorFull=Blue2Full
		Light23b.Intensity=Blue2I
		Light23b1.Color=Blue2
		Light23b1.ColorFull=Blue2Full
		Light23b1.Intensity=Blue2I
		Light23b2.Color=Blue2
		Light23b2.ColorFull=Blue2Full
		Light23b2.Intensity=Blue2I

		Light22b.Color=Yellow2
		Light22b.ColorFull=Yellow2Full
		Light22b.Intensity=Yellow2I
		Light22b1.Color=Yellow2
		Light22b1.ColorFull=Yellow2Full
		Light22b1.Intensity=Yellow2I
		Light22b2.Color=Yellow2
		Light22b2.ColorFull=Yellow2Full
		Light22b2.Intensity=Yellow2I

		Light21b.Color=Orange
		Light21b.ColorFull=OrangeFull
		Light21b.Intensity=OrangeI
		Light21b1.Color=Orange
		Light21b1.ColorFull=OrangeFull
		Light21b1.Intensity=OrangeI
		Light21b2.Color=Orange
		Light21b2.ColorFull=OrangeFull
		Light21b2.Intensity=OrangeI

		Light20b.Color=Red
		Light20b.ColorFull=RedFull
		Light20b.Intensity=RedI
		Light20b1.Color=Red
		Light20b1.ColorFull=RedFull
		Light20b1.Intensity=RedI
		Light20b2.Color=Red
		Light20b2.ColorFull=RedFull
		Light20b2.Intensity=RedI

		Light19b.Color=White
		Light19b.ColorFull=WhiteFull
		Light19b.Intensity=WhiteI
		Light19b1.Color=White
		Light19b1.ColorFull=WhiteFull
		Light19b1.Intensity=WhiteI
		Light19b2.Color=White
		Light19b2.ColorFull=WhiteFull
		Light19b2.Intensity=WhiteI
End If

If HottieMod = 0 Then
	HottieModType = Int(Rnd*10)+1
Else
	HottieModType = HottieMod
End If

If HottieModType = 1 Then
	Face.image = "BOPHead"
End If

If HottieModType = 2 Then
	Face.image = "_Ashley"
End If

If HottieModType = 3 Then
	Face.image = "_Brooke"
End If

If HottieModType = 4 Then
	Face.image = "_Kelly"
End If

If HottieModType = 5 Then
	Face.image = "_Brunette"
End If

If HottieModType = 6 Then
	Face.image = "_Meagan"
End If

If HottieModType = 7 Then
	Face.image = "_Shelly"
End If

If HottieModType = 8 Then
	Face.image = "_Rachel"
End If

If HottieModType = 9 Then
	Face.image = "_Valerie"
End If

If HottieModType = 10 Then
	Face.image = "_Harley"
End If

End Sub




'Flipper Rubbers Color Mod


If Choosebats = 0 Then
	ChooseBatsType = Int(Rnd*3)+1
Else
	ChoosebatsType = Choosebats
End If

If ChoosebatsType = 1 Then
			'batleft.visible = 1 : batright.visible = 1 : LeftFlipper.visible = 0 : RightFlipper.visible = 0
			'batleftshadow.visible = 1 : batrightshadow.visible = 1 : GraphicsTimer.enabled = True
			batleft.image = "flipper_white_black" : batright.image = "flipper_white_black"
End If

If ChoosebatsType = 2 Then
			'batleft.visible = 1 : batright.visible = 1 : LeftFlipper.visible = 0 : RightFlipper.visible = 0
			'batleftshadow.visible = 1 : batrightshadow.visible = 1 : GraphicsTimer.enabled = True
			batleft.image = "flipper_white_red" : batright.image = "flipper_white_red"
End If

If ChoosebatsType = 3 Then
			'batleft.visible = 1 : batright.visible = 1 : LeftFlipper.visible = 0 : RightFlipper.visible = 0
			'batleftshadow.visible = 1 : batrightshadow.visible = 1 : GraphicsTimer.enabled = True
			batleft.image = "flipper_white_blue" : batright.image = "flipper_white_blue"
End If


Sub GraphicsTimer_Timer()
		'If ChooseBats > 0 Then
		' *** move primitive bats ***
		batleft.objrotz = LeftFlipper.CurrentAngle + 1
		batleftshadow.objrotz = batleft.objrotz
		batright.objrotz = RightFlipper.CurrentAngle - 1
		batrightshadow.objrotz  = batright.objrotz
	'End If
End Sub




'Example Usage (Needs to be piggybacked off of a sub that handles the fading, such as nModFlash)
'            Lamp Number     Object      Image Prefix    Number of images (in this case 0 is full off, 13 is full on)
'ModFlashObjm 420,           F20P,       "DomeYellow_",  13

Sub ModFlashObjm(nr, object, imgseq, steps)    'Primitive texture image sequence    'Modified for SolModCallbacks
   dim x, x2, fadex
    Select Case FadingLevel(nr)
        Case 3, 4, 5, 6 'off
           FadeX = ScaleLights(FlashLevel(nr),0 )    'mod for solmodcallbacks
           for x = 0 to steps-1
                if FadeX <= ((x/steps) + ((1/steps)/2 )) then
                    x2 = x
'                    tb3.text = "on " & x & vbnewline & ((x/steps) + ((1/steps)/2 )) & " =? " & x2
                   exit for
                end if
            next
            if    FadeX >= 1-(1/steps) then x2 = steps ': tb3.text = "fullon"
       '    if    FlashLevel(nr) <= (1/steps) then x2 = 0 : tb3.text = "fulloff"
'            tb.text = "flashlevel: " & FlashLevel(nr) & vbnewline & "stepper: " & x2 & vbnewline & "flasherimg: " & x2
'            tb.text = FlashLevel(nr) & vbnewline & imgseq & x2
           object.image = imgseq & x2
    End Select
End Sub

' _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _
'(_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_)

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
  Vol = Csng(BallVel(ball) ^2 / 2000)
End Function

Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
  Pitch = BallVel(ball) * 20
End Function

Function BallVel(ball) 'Calculates the ball speed
  BallVel = INT(SQR((ball.VelX ^2) + (ball.VelY ^2) ) )
End Function

'*****************************************
'    JP's VP10 Rolling Sounds
'*****************************************

Const tnob = 8 ' total number of balls
ReDim rolling(tnob)
InitRolling

Sub InitRolling
    Dim i
    For i = 0 to tnob
        rolling(i) = False
    Next
End Sub

Sub RollingSoundTimer_Timer()
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
Sub Table_exit()
  Controller.Pause = False
  Controller.Stop
End Sub


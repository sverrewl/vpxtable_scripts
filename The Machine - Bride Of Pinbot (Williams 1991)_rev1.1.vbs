
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


'The Machine (Bride of Pinbot) Williams 1991 Rev1.1 for VP10 (Requires VP10.2 or greater to play)

'***The very talented BOP development team.***
'Original VP10 beta by "Unclewilly and completed by "wrd1972"
'Additional scripting assistance by "cyberpez", "Rothbauerw".
'Clear ramps and wire ramp primitives by "Flupper"
'Additional 3D work by "Dark" and "Cyberpez"
'Space shuttle toy and sidewall graphics by "Cyberpez"
'Face rotation scripting by "KieferSkunk/Dorsola"
'High poly helmet and re-texture by "Dark"
'Original bride helmet, plastics and playfield redraw by "Jc144"
'Desktop view, scoring dials aand background by "32Assassin"
'Playfield lighting by "Haunt Freaks"
'Flashers by "wrd1972"
'Physics by "Clarkkent"
'DOF and controller by "Arngrim"
'***Many...many thanks to you all for the tremenedous efforts and hard work for making this awesome table a reality. I just cant thank you all enough.*** 


Option Explicit
Randomize

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Dim cGameName, ShipMode, cheaterpost, sidewalls

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

'<<<Show shuttle ramp ship toy>>>
'0 No ship
'1 Show Ship

ShipMode = 0


'<<<Set sidewalls style here>>>
'0=Random Sidewall
'1=Wood walls
'2=Starfield walls 
'3=BOP CabArt Style

sidewalls =  3   




'<<<Add drain post here>>>
'0 No drain post
'1 Show drain post

cheaterpost = 0  



' _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ 
'(_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_)

'RomNames
 cGameName = "bop_l7"

Const BallSize = 25  'radius
Const ballmass = 1.05

Dim DesktopMode: DesktopMode = Table.ShowDT

If DesktopMode = True Then 'Show Desktop components
	leftrail.visible=1
	rightrail.visible=1
	Face1.visible=1
	face.visible=0
	frontlockbar.visible=1
	rearbar.visible=1
	pSidewall.Z = -100
	F64c. visible = 1
	F64d. visible = 0
	FaceguidesDT. visible = 1
	FaceguidesFS. visible = 0
	l46_DT. visible = 1
	l47_DT. visible = 1
	l46_FS. visible = 0
	l47_FS. visible = 0
	l48a_DT. visible = 1
	l48a_FS. visible = 0
	f_46a_fs. visible = 0
	f_47a_fs. visible = 0

Else
	leftrail.visible=0
	rightrail.visible=0
	Face1.visible=0
	Face.visible=1
	frontlockbar.visible=0
	rearbar.visible=0
	pSidewall.Z = -50
	F64c. visible = 0
	F64d. visible = 1
	FaceguidesDT. visible = 0
	FaceguidesFS. visible = 1
	l46_DT. visible = 0
	l47_DT. visible = 0
	l46_FS. visible = 1
	l47_FS. visible = 1
	l48a_DT. visible = 0
	l48a_FS. visible = 1
	f_46a_fs. visible = 1
	f_47a_fs. visible = 1



End if

'LoadCoreVBS  

Sub LoadCoreVBS
     On Error Resume Next
     ExecuteGlobal GetTextFile("core.vbs")
     If Err Then MsgBox "Can't open core.vbs"
     On Error Goto 0
End Sub

LoadVPM "01560000", "WPC.VBS", 3.26

   'Variables
Dim bsSaucer, bsLEye, bsREye, bsMouth, bsSS, mFace, xx, bump1, bump2, bump3
Dim mechHead, headAngle, prevHeadAngle, currentFace
Dim MaxBalls 

MaxBalls=3	

Const UseSolenoids = 1
Const UseLamps = 0
Const UseSync = 1
Const HandleMech = 0 

Set GICallback2 = GetRef("UpdateGI")

'Standard Sounds
Const SSolenoidOn = "Solenoid"
Const SSolenoidOff = ""
Const SCoin = "CoinIn"
 
' Setup the lightning according to the nightday slider
Dim ii
for each ii in aGiTopLights: ii.intensity = ii.intensity + (100 - table.nightday)/10: next
for each ii in aGiBottomLights: ii.intensity = ii.intensity + (100 - table.nightday)/10: next
for each ii in aLights: ii.intensity = ii.intensity + (100 - table.nightday)/8: next
for each ii in aFlashers:ii.opacity=ii.opacity + (100 - table.nightday)^2:next


'Table Init
Sub Table_Init
	vpmInit Me
	With Controller
		.GameName =  cGameName
		If Err Then MsgBox "Can't start Game " & cGameName & vbNewLine & Err.Description:Exit Sub
		.SplashInfoLine = "The Machine BOP, Williams 1991" & vbNewLine & "by unclewilly vp9"
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

'Nudging
	vpmNudge.TiltSwitch=14
    vpmNudge.Sensitivity=1
    vpmNudge.TiltObj=Array(Bumper1b,Bumper2b,Bumper3b,LeftSlingshot,RightSlingshot,UpperSlingShot)

          Set bsLEye = New cvpmBallStack
          With bsLEye
             .InitSaucer sw63,63, 180, 5
            .InitExitSnd SoundFX("Scoopexit",DOFContactors), SoundFX("Solenoid",DOFContactors)
  		   .InitAddSnd "kicker_enter"
          End With
 
          Set bsREye = New cvpmBallStack
          With bsREye
             .InitSaucer sw64,64, 180, 5
             .InitExitSnd SoundFX("Scoopexit",DOFContactors), SoundFX("Solenoid",DOFContactors)
  			.InitAddSnd "kicker_enter"
          End With
 
          Set bsMouth = New cvpmBallStack
          With bsMouth
             .InitSaucer sw65,65, 180, 5
             .InitExitSnd SoundFX("Scoopexit",DOFContactors), SoundFX("Solenoid",DOFContactors)
  			.InitAddSnd "kicker_enter"
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
     If vpmKeyDown(keycode) Then Exit Sub 
   'If keycode = 30 then TestKicker.CreateBall:TestKicker.Kick 180, 1
End Sub

Sub Table_KeyUp(ByVal keycode)

 	 If keycode = plungerkey then plunger.Fire:PlaySound "Plunger2"
	 'If Keycode = RightFlipperKey then SolRFlipper 0
	 'If Keycode = LeftFlipperKey then SolLFlipper 0
	If vpmKeyUp(keycode) Then Exit Sub

End Sub

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
 	'Flashers
       SolCallback(17) = "Flash17" 'Billions
       SolCallback(18) = "SetFlash 118," 'Left ramp
       SolCallback(19) = "Flash19" 'jackpot
       SolCallback(20) = "Flash20" 'SkillShot
       SolCallback(21) = "SetFlash 121," 'Left Helmet
       SolCallback(22) = "SetFlash 122," 'Right Helmet
       SolCallback(23) = "SetFlash 123," 'Jets Enter
       SolCallback(24) = "SetFlash 124," 'Left Loop
       'SolCallback(28) = "SolMotor"
      ' SolCallback(27) = "SolRelay"


Sub SolKickout(enabled)
	If Enabled then
		sw46k.kick 22.5,55
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
 	If Enabled then	BL.IsDropped=1:BLP.transY = -90:BL.TimerEnabled = 1
   End Sub

	Sub BL_Timer()
		BL.IsDropped=0:BLP.transY = 0:BL.TimerEnabled = 0
	End Sub

'Flashers
Sub Flash17(enabled)
	Setlamp 117, enabled
End Sub

Sub Flash19(enabled)
	Setlamp 119, enabled
End Sub

Sub Flash20(enabled)
	Setlamp 120, enabled
End Sub

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
'	smlGateP.RotZ = (smlGate.currentangle)
	SpinnerT1.RotZ = -(sw51.currentangle)
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
	LogoL.ObjRotz = LeftFlipper.Currentangle+180	
	LogoR.ObjRotz = RightFlipper.Currentangle+180
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
    Face1.ObjRotY = headAngle
    FaceGuidesFS.ObjRotY = headAngle
	pFaceDiverter.ObjRotY = headAngle
	pFaceDiverterPegs.ObjRotY = headAngle
    FaceGuidesDT.ObjRotY = headAngle

  ' Determine which face is up, if any
    If headIsNear(0) OR headIsNear(360) Then
	  ShowEyeBulbs(True)
      currentFace = 1
      DOF 102, DOFOff
    ElseIf headIsNear(90) Then
      ShowEyeBulbs(True)
	  currentFace = 2
	  DOF 102, DOFoff
    ElseIf headIsNear(180) Then
      ShowEyeBulbs(True)
      currentFace = 3
      DOF 102, DOFOff
    ElseIf headIsNear(270) Then
      ShowEyeBulbs(True)
      currentFace = 4
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
			L46_dt.intensity = 15
			L47_dt.intensity = 15
			L46_fs.intensity = 4
			L47_fs.intensity = 4
			f_46a_fs.opacity = 500
			f_47a_fs.opacity = 500
	  Else
			L46_dt.intensity = 0
			L47_dt.intensity = 0
			L46_fs.intensity = 0
			L47_fs.intensity = 0
			f_46a_fs.opacity = 0
			f_47a_fs.opacity = 0
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
   Sub SSLaunch_Hit():bsSS.AddBall 0:End Sub

Dim DivPOS, SideWallType

'''''''' SetOptions
Sub SetOptions()

'ShipMod
	If ShipMode = 1 Then
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
	
End Sub

  
   'Bumpers
      Sub Bumper1b_Hit:vpmTimer.PulseSw 53:PlaySound SoundFx("fx_bumper",DOFContactors):End Sub 'bump1 = 1:Me.TimerEnabled = 1
     
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
 
      Sub Bumper2b_Hit:vpmTimer.PulseSw 55:PlaySound SoundFx("fx_bumper",DOFContactors):End Sub 'bump2 = 1:Me.TimerEnabled = 1
     
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
 
      Sub Bumper3b_Hit:vpmTimer.PulseSw 54:PlaySound SoundFx("fx_bumper",DOFContactors):End Sub 'bump3 = 1:Me.TimerEnabled = 1
     
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


 'Rollovers

 'FlipperLanes and Plunger
   Sub sw15_Hit:Controller.Switch(15) = 1:PlaySound "rollover":End Sub
   Sub sw15_UnHit:Controller.Switch(15) = 0:End Sub
   Sub sw16_Hit:Controller.Switch(16) = 1:PlaySound "rollover":End Sub
   Sub sw16_UnHit:Controller.Switch(16) = 0:End Sub
   Sub sw17_Hit:Controller.Switch(17) = 1:PlaySound "rollover":End Sub
   Sub sw17_UnHit:Controller.Switch(17) = 0:End Sub
   Sub sw18_Hit:Controller.Switch(18) = 1:PlaySound "rollover":End Sub
   Sub sw18_UnHit:Controller.Switch(18) = 0:End Sub
   Sub sw52_Hit:Controller.Switch(52) = 1:PlaySound "rollover":End Sub
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


'Sounds
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


'MISC. Sound effects

Sub ssrubbers_Hit(): playsound"fx_bump" : End Sub
Sub Subway_Hit(): playsound"fx_metalrolling" : End Sub
Sub UPFdrop_Hit: playsound"ball_bounce" : End Sub
Sub xrampdrop_Hit: playsound"ball_bounce" : End Sub
Sub lrampdrop_Hit: playsound"ball_bounce" : End Sub
Sub Wall50_hit():PlaySound "Bump":End Sub
Sub Wall51_hit():PlaySound "Bump":End Sub
Sub Wall52_hit():PlaySound "Bump":End Sub
Sub Wall53_hit():PlaySound "Bump":End Sub
Sub Wall57_hit():PlaySound "Bump":End Sub
Sub BL_hit():PlaySound "wirerampdrop":End Sub
Sub left_gate_Hit():PlaySound "fx_gate2":End Sub
Sub right_gate_Hit():if gate3.currentangle < 90 then:PlaySound "fx_gate2":end if:End Sub


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


Sub RandomSoundRubber()
	Select Case Int(Rnd*3)+1
		Case 1 : PlaySound "rubber_hit_1"
		Case 2 : PlaySound "rubber_hit_2"
		Case 3 : PlaySound "rubber_hit_3"
	End Select
End Sub

Sub RandomSoundRubberLowVolume()
	Select Case Int(Rnd*3)+1
		Case 1 : PlaySound "rubber_hit_1_low"
		Case 2 : PlaySound "rubber_hit_2_low"
		Case 3 : PlaySound "rubber_hit_3_low"
	End Select
End Sub
 
'Wood Sounds
Sub Woods_Hit(idx)
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  	if finalspeed > 2 and activeball.velx > 1 then PlaySound "woodhit"
End sub
 
 'Plastic Sounds
Sub Plastics_Hit(idx)
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  	if finalspeed > 2 then PlaySound "plastichit"
End sub
 
'Pins


Sub LeftFlipper_Collide(parm)
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 4 then 
		RandomSoundFlipper()
	Else 
 		RandomSoundFlipperLowVolume()
 	End If
End Sub

Sub RightFlipper_Collide(parm)
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 4 then 
		RandomSoundFlipper()
	Else 
 		RandomSoundFlipperLowVolume()
 	End If
End Sub

Sub RandomSoundFlipper()
	Select Case Int(Rnd*3)+1
		Case 1 : PlaySound "flip_hit_1"
		Case 2 : PlaySound "flip_hit_2"
		Case 3 : PlaySound "flip_hit_3"
	End Select
End Sub

Sub RandomSoundFlipperLowVolume()
	Select Case Int(Rnd*3)+1
		Case 1 : PlaySound "flip_hit_1_low"
		Case 2 : PlaySound "flip_hit_2_low"
		Case 3 : PlaySound "flip_hit_3_low"
	End Select
End Sub

'*****************
' GI2 Illumination (Variable Lights 1-8 and off - Williams / Bally)
' 8 Step Sub
'*****************
'*****************
' GI2 Illumination (Variable Lights 1-8 and off - Williams / Bally)
' 8 Step Sub
'*****************

'***********
' Update GI
'***********
'*****Init gi

Dim gistep
gistep = 1 / 8
UpdateGI 0,0:UpdateGI 1,0:UpdateGI 2,0:UpdateGI 3,0:UpdateGI 4,0
Sub UpdateGI(no, step)
    If step = 7 then exit sub    '0 OR step = 
    Select Case no
        Case 0  'backglass
''            For each xx in BlueGIHalo:xx.alpha = gistep * step:next
''            GISL.alpha = gistep * step

        Case 1   'helmet
'			HelGI.IntensityScale = cint(gistep * step)
'			HelGI1.IntensityScale = cint(gistep * step)
        Case 2   'rear pf
			For each xx in aGiTopLights:xx.intensityscale = gistep * step:next
				 'Playsound "fx_relay_on"
        Case 3   'backglass
'
        Case 4  'front pf
			For each xx in aGiBottomLights:xx.IntensityScale = gistep * step:next
                 'Playsound "fx_relay_on"
    End Select
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

AllLampsOff()
LampTimer.Interval = 30
 'lamp fading speed
LampTimer.Enabled = 1

FlashInit()
FlasherTimer.Interval = 10 'flash fading speed
FlasherTimer.Enabled = 1

' Lamp & Flasher Timers

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
	'UpdateLeds
End Sub
 
 Sub UpdateLamps

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
    flash 37, f37
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


  	NFadeLFOm 41
	nFadel 41, Light19b
    flash 41, f19a


  	NFadeLFOm 42
	nFadel 42, Light20b
    flash 42, f20a

  	NFadeLFOm 43
	nFadel 43, Light21b
    flash 43, f21a

  	NFadeLFOm 44
	nFadel 44, Light22b
    flash 44, f22a

  	NFadeLFOm 45
	NFadeL 45, Light23b
    flash 45, f23a

	NFadeLm 46, l46_dt
	NFadeLm 46, l46_fs
    Flash 46, F_46a_fs

	NFadeLm 47, l47_dt
	NFadeLm 47, l47_fs
   Flash 47, F_47a_fs

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
	NFadeLFOm 64
	nFadeL 65, l65
	nFadeL 66, l66
	nFadeL 67, l67
	nFadeL 68, l68





'Helmet

  	nFadeLm 91, l91
    Flash 91, f91


  	nFadeLm 92, l92
    Flash 92, f92

  	nFadeLm 93, l93
    Flash 93, f93

  	nFadeLm 94, l94
    Flash 94, f94

  	nFadeLm 95, l95
    Flash 95, f95

  	nFadeLm 96, l96
    Flash 96, f96

  	nFadeLm 97, l97
    Flash 97, f97

  	nFadeLm 98, l98
    Flash 98, f98

  	nFadeLm 101, l101
    Flash 101, f101

  	nFadeLm 102, l102
    Flash 102, f102

  	nFadeLm 103, l103
    Flash 103, f103


  	nFadeLm 104, l104
    Flash 104, f104

  	nFadeLm 105, l105
    Flash 105, f105

  	nFadeLm 106, l106
    Flash 106, f106

  	nFadeLm 107, l107
    Flash 107, f107

  	nFadeLm 108, l108
    Flash 108, f108




	'***Flashers
	NFadeL 117, f17
	NFadeL 119, f19

  	NFadeLFOm 120
	NFadelm 120, L120
	'NFadel 120, L120a
	Flash 120, F120


 End Sub
 
Sub FlasherTimer_Timer()

	Flash 41, l41b
	Flash 42, l42b
	Flash 43, l43b
	Flash 44, l44b
	Flash 45, l45b
'	Flash 46, f48b
'	Flash 47, f48c

'	nfadelm 48, l48
	nfadelm 48, l48a_dt
	nfadelm 48, l48a_fs
	nfadelm 48, l48b
	Flash 48, f48a

	Flashm 86, f86
	Flash 86, f86b
	Flashm 87, f87
	Flash 87, f87b
	Flash 88, f88


'	'****Flashers
	Flashm 64, F64a
	Flashm 64, F64b
	Flashm 64, F64c
	Flash 64, F64d
	Flash 118, f18
'	Flash 120, f120
	Flashm 121, F121a
	Flash 121, F121b
	Flashm 122, F122a
	Flash 122, F122b
	Flashm 123, F123a
	Flash 123, F123b
	If ShipMode = 1 then
		Flashm 124, F124c
		Flashm 124, F124d
		Flashm 124, F124e
		Flashm 124, F124f
		Flashm 124, F124g
		Flashm 124, F124h
	End If
	Flashm 124, F124a
	Flash 124, F124b
End Sub

' div lamp subs

Sub AllLampsOff()
    Dim x
    For x = 0 to 200
        LampState(x) = 0
        FadingLevel(x) = 4
    Next

'UpdateLamps:UpdateLamps:Updatelamps
End Sub

Sub SetLamp(nr, value)
    If value <> LampState(nr) Then
        LampState(nr) = abs(value)
        FadingLevel(nr) = abs(value) + 4
    End If
End Sub

' div flasher subs

Sub FlashInit
    Dim i
    For i = 0 to 200
        FlashState(i) = 0
        FlashLevel(i) = 0
    Next

    FlashSpeedUp = .5   ' fast speed when turning on the flasher
    FlashSpeedDown = .1 ' slow speed when turning off the flasher, gives a smooth fading
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
            Object.intensityscale = FlashLevel(nr)
        Case 1 ' on
            FlashLevel(nr) = FlashLevel(nr) + FlashSpeedUp
            If FlashLevel(nr) > 1 Then
                FlashLevel(nr) = 1
                FlashState(nr) = -2 'completely on
            End if
            Object.intensityscale = FlashLevel(nr)
    End Select
End Sub

Sub FlashP(nr, object, a, b, c, d)
    Select Case FlashState(nr)
        Case 0 'off
            FlashLevel(nr) = FlashLevel(nr) - FlashSpeedDown
            If FlashLevel(nr) < 0 Then
                FlashLevel(nr) = 0
                FlashState(nr) = -1 'completely off
            End if
			Select case FlashLevel(nr)
				Case 0: object.image = d
				Case .34: object.image = c
				Case .66: object.image = b
				Case 1: object.image = a
			end select
        Case 1 ' on
            FlashLevel(nr) = FlashLevel(nr) + FlashSpeedUp
            If FlashLevel(nr) > 1 Then
                FlashLevel(nr) = 1
                FlashState(nr) = -2 'completely on
            End if
			Select case FlashLevel(nr)
				Case 0: object.image = d
				Case .34: object.image = c
				Case .66: object.image = b
				Case 1: object.image = a
			end select
    End Select
End Sub

Sub FlashC(nr, object,object2,a,b)

    Select Case FlashState(nr)
        Case 0 'off
            FlashLevel(nr) = FlashLevel(nr) - FlashSpeedDown
            If FlashLevel(nr) < 0 Then
                FlashLevel(nr) = 0
                FlashState(nr) = -1 'completely off
            End if
			If headAngle = a then 
				object.alpha = FlashLevel(nr)
				object2.alpha = 0
			end if
			If headAngle = b then 
				object.alpha = 0
				object2.alpha = FlashLevel(nr)
			End If 
			If headAngle <> a and headAngle <> b then 
				object.alpha = 0
				object2.alpha = 0
			End If         
        Case 1 ' on
            FlashLevel(nr) = FlashLevel(nr) + FlashSpeedUp
            If FlashLevel(nr) > 255 Then
                FlashLevel(nr) = 255
                FlashState(nr) = -2 'completely on
            End if
			If headAngle = a then 
				object.alpha = FlashLevel(nr)
				object2.alpha = 0
			end if
			If headAngle = b then 
				object.alpha = 0
				object2.alpha = FlashLevel(nr)
			End If
			If headAngle <> a and headAngle <> b then 
				object.alpha = 0
				object2.alpha = 0
			End If
    End Select

End Sub

Sub FlashD(nr, object,object2,a,b,c,d)
    Select Case FlashState(nr)
        Case 0 'off
            FlashLevel(nr) = FlashLevel(nr) - FlashSpeedDown
            If FlashLevel(nr) < 0 Then
                FlashLevel(nr) = 0
                FlashState(nr) = -1 'completely off
            End if
			If headAngle = b then 
				object.alpha = FlashLevel(nr)
				object2.alpha = 0
			end if
			If headAngle = a or headAngle = c or headAngle = d then 
				object.alpha = 0
				object2.alpha = FlashLevel(nr)
			End If 
			If headAngle <> a and headAngle <> b and headAngle <> c and headAngle <> d then 
				object.alpha = 0
				object2.alpha = 0
			End If        
        Case 1 ' on
            FlashLevel(nr) = FlashLevel(nr) + FlashSpeedUp
            If FlashLevel(nr) > 255 Then
                FlashLevel(nr) = 255
                FlashState(nr) = -2 'completely on
            End if
			If headAngle = b then 
				object.alpha = FlashLevel(nr)
				object2.alpha = 0
			end if
			If headAngle = a or headAngle = c or headAngle = d then 
				object.alpha = 0
				object2.alpha = FlashLevel(nr)
			End If
			If headAngle <> a and headAngle <> b and headAngle <> c and headAngle <> d then 
				object.alpha = 0
				object2.alpha = 0
			End If
    End Select

End Sub

Sub Flashm(nr, object) 'multiple flashers, it doesn't change the flashstate
    Select Case FlashState(nr)
        Case 0         'off
            Object.intensityscale = FlashLevel(nr)
        Case 1         ' on
            Object.intensityscale = FlashLevel(nr)
    End Select
End Sub

'Special
Sub FadeFlames(nr)
    Select Case FadingLevel(nr)
        Case 2:FlamesFade = 0:FadingLevel(nr) = 0 'Off
        Case 3:FlamesFade = 1:FadingLevel(nr) = 2 'fading...
        Case 4:FlamesFade = 2:FadingLevel(nr) = 3 'fading...
        Case 5:FlamesFade = 3:FadingLevel(nr) = 1 'ON
    End Select
End Sub

Sub NFadeLFO(nr)
    Select Case FadingLevel(nr)
        Case 4:SetFlash nr, 0:FadingLevel(nr) = 0
        Case 5:SetFlash nr, 1:FadingLevel(nr) = 1
    End Select
End Sub

Sub NFadeLFOm(nr)
    Select Case FadingLevel(nr)
        Case 4:SetFlash nr, 0
        Case 5:SetFlash nr, 1
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

Sub FadeL(nr, a)
    Select Case FadingLevel(nr)
        Case 2:a.intensityscale = 0:FadingLevel(nr) = 0
        Case 3:a.intensityscale = .34:FadingLevel(nr) = 2
        Case 4:a.intensityscale = .66:FadingLevel(nr) = 3
        Case 5:a.intensityscale = 1:FadingLevel(nr) = 1
    End Select
End Sub

Sub FadeLm(nr, a)
    Select Case FadingLevel(nr)
        Case 2:a.intensityscale = 0
        Case 3:a.intensityscale = .34
        Case 4:a.intensityscale = .66
        Case 5:a.intensityscale = 1
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
    PlaySound SoundFX("fx_slingshot",DOFContactors),0,1,-0.05,0.05
    USling.Visible = 0
    USling1.Visible = 1
    slingu.TransZ = -20
    UStep = 0
    UpperSlingShot.TimerEnabled = 1
	vpmTimer.PulseSw 55
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
    PlaySound SoundFX("fx_slingshot",DOFContactors),0,1,-0.05,0.05:vpmTimer.PulseSw 57
    LSling.Visible = 0
    LSling3.Visible = 1
    sling2.TransZ = -20
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
	vpmTimer.PulseSw 55
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
    PlaySound SoundFX("fx_slingshot",DOFContactors), 0, 1, 0.05, 0.05:vpmTimer.PulseSw 58
    RSling.Visible = 0
    RSling3.Visible = 1
    sling1.TransZ = -20
    RStep = 0
    RightSlingShot.TimerEnabled = 1
	vpmTimer.PulseSw 56
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

Sub WireStart_Hit():PlaySound "fx_metalrolling":End Sub



' *********************************************************************
'                      Supporting Ball & Sound Functions
' *********************************************************************

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed, decrease the default 2000 to hear a louder rolling ball sound
    Vol = Csng(BallVel(ball) ^2 / 1100)
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "table1" is the name of the table
    Dim tmp
    tmp = ball.x * 2 / table.width-1
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
'    JP's VP10 Rolling Sounds
'*****************************************

Const tnob = 8 ' total number of balls
ReDim rolling(tnob)
InitRolling

Sub testkick()
test.CreateSizedBallWithMass BallSize, BallMass
test.kick 0,60
End Sub

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
    PlaySound("collide3"), 0, Csng(velocity) ^2 / 2000, Pan(ball1), 0, Pitch(ball1), 0, 0
End Sub

Sub Gates_Hit (idx)
	PlaySound "gate", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
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

Sub Trigger1_Hit():PlaySound "fx_lr5":End Sub
Sub Trigger2_Hit():PlaySound "fx_lr5":End Sub
Sub Trigger3_Hit():PlaySound "fx_lr5":End Sub
Sub Trigger4_Hit():PlaySound "fx_lr5":End Sub
Sub Trigger5_Hit():PlaySound "fx_lr1":End Sub
Sub Trigger6_Hit():PlaySound "fx_lr5" End Sub
Sub Trigger7_Hit():PlaySound "fx_lr5":End Sub
Sub Trigger8_Hit():PlaySound "fx_lr2":End Sub
Sub Trigger9_Hit():PlaySound "fx_lr6":End Sub
Sub Trigger10_Hit():PlaySound "fx_lr6":End Sub

Sub RubberslowerPFlargepost_Hit(idx)
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 3 then 
		PlaySound "fx_rubber2", 0, 2*Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
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
'	Kicker1active = 1
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


















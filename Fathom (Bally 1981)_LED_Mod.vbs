Option Explicit
Randomize

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Const cGameName="Fathom",UseSolenoids=1,UseLamps=0,UseGI=0,SSolenoidOn="SolOn",SSolenoidOff="SolOff", SCoin="coin"

LoadVPM "01130100", "Bally.VBS", 3.21
Dim DesktopMode: DesktopMode = Table1.ShowDT

If DesktopMode = True Then 'Show Desktop components
Ramp16.visible=1
Ramp15.visible=1
Primitive13.visible=1
Else
Ramp16.visible=0
Ramp15.visible=0
Primitive13.visible=0
End if

'*************************************************************
'Solenoid Call backs
'**********************************************************************************************************
 SolCallback(1)  = "dtT.SolDropUp"
 SolCallback(2)  = "dtL.SolDropUp"
 SolCallback(3)  = "dtM.SolDropUp"
 SolCallback(4)  = "dtR.SolDropUp"
 SolCallback(6)	 =  "vpmSolSound SoundFX(""Knocker"",DOFKnocker),"
 SolCallback(7)  = "bsTrough.SolOut"
 SolCallback(13) = "bsTopEject.SolOut"
 SolCallback(14) = "bsRightEject.SolOut"
 SolCallback(15) = "dtT.SolHit 3,"
 SolCallback(25) = "dtR.SolHit 1,"
 SolCallback(26) = "dtR.SolHit 2,"
 SolCallback(27) = "dtR.SolHit 3,"
 SolCallback(37) = "dtT.SolHit 1,"
 SolCallback(38) = "dtT.SolHit 2,"

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
     If Enabled Then
         PlaySound SoundFX("fx_Flipperup",DOFContactors):LeftFlipper.RotateToEnd
     Else
         PlaySound SoundFX("fx_Flipperdown",DOFContactors):LeftFlipper.RotateToStart
     End If
  End Sub
  
Sub SolRFlipper(Enabled)
     If Enabled Then
         PlaySound SoundFX("fx_Flipperup",DOFContactors):RightFlipper.RotateToEnd:RightFlipper1.RotateToEnd
     Else
         PlaySound SoundFX("fx_Flipperdown",DOFContactors):RightFlipper.RotateToStart:RightFlipper1.RotateToStart
     End If
End Sub

'**********************************************************************************************************

'Solenoid Controlled toys
'**********************************************************************************************************

'*****GI Lights On
dim xx
For each xx in GI:xx.State = 1: Next

'Primitive Gate
Sub FlipperTimer_Timer
	PrimGate2.Rotz = Gate7.Currentangle 
 End Sub

   Sub SolTimer_Timer
 	Dim ChgSol, tmp, ii, CBoard, solnum
	ChgSol  = Controller.ChangedSolenoids
	If Not IsEmpty(ChgSol) Then
 	CBoard = Controller.Lamp(47)
		For ii = 0 To UBound(ChgSol)
 			solnum = ChgSol(ii, 0)
 			If solnum <= 14 and CBoard Then solnum = solnum + 24
 			tmp = Solcallback(solnum)
			If tmp <> "" Then Execute tmp & vpmTrueFalse(ChgSol(ii, 1)+1)
		Next
	End If
   End Sub

'**********************************************************************************************************

'Initiate Table
'**********************************************************************************************************
Dim bsTrough, bsRightEject, bsTopEject, dtT, dtL, dtM, dtR
Sub Table1_Init
	vpmInit Me
	On Error Resume Next
		With Controller
		.GameName = cGameName
		If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
		.SplashInfoLine = "Fathom (Bally)"&chr(13)&"You Suck"
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

	vpmNudge.TiltSwitch  = 15
	vpmNudge.Sensitivity = 5
	vpmNudge.Tiltobj = Array(LeftSlingshot,RightSlingshot,Bumper1,Bumper2,Bumper3)

	' Trough
	Set bsTrough = New cvpmBallStack
		bsTrough.Initsw 0,1,2,3,0,0,0,0
		bsTrough.InitKick BallRelease,55,10
		bsTrough.InitExitSnd SoundFX("ballrelease",DOFContactors), SoundFX("Solenoid",DOFContactors)
		bsTrough.Balls = 3

 	' Top Saucer
	Set bsTopEject = New cvpmBallStack
		bsTopEject.InitSaucer sw4, 4, 275, 13
		bsTopEject.InitExitSnd SoundFX("Popper",DOFContactors), SoundFX("Solenoid",DOFContactors)

 	' Right Saucer
	Set bsRightEject = New cvpmBallStack
		bsRightEject.InitSaucer sw5, 5, 180, 9
		bsRightEject.InitExitSnd SoundFX("Popper",DOFContactors), SoundFX("Solenoid",DOFContactors)

 	Set dtL=New cvpmDropTarget
		dtL.InitDrop Array(sw27,sw28,sw29,sw30,sw31,sw32),Array(27,28,29,30,31,32)
		dtL.InitSnd SoundFX("DTDrop",DOFContactors),SoundFX("DTReset",DOFContactors)

 	Set dtM=New cvpmDropTarget
		dtM.InitDrop Array(sw33,sw34,sw35),Array(33,34,35)
		dtM.InitSnd SoundFX("DTDrop",DOFContactors),SoundFX("DTReset",DOFContactors)

 	Set dtT=New cvpmDropTarget
		dtT.InitDrop Array(Array(sw44,sw44a),Array(sw43,sw43a),Array(sw42,sw42a)),Array(44,43,42)
		dtT.InitSnd SoundFX("DTDrop",DOFContactors),SoundFX("DTReset",DOFContactors)

 	Set dtR=New cvpmDropTarget
		dtR.InitDrop Array(sw48,sw47,sw46),Array(48,47,46)
		dtR.InitSnd SoundFX("DTDrop",DOFContactors),SoundFX("DTReset",DOFContactors)

End Sub

'**********************************************************************************************************
'Plunger code
'**********************************************************************************************************

Sub Table1_KeyDown(ByVal KeyCode)
	If KeyDownHandler(keycode) Then Exit Sub
	If keycode = PlungerKey Then Plunger.Pullback:playsound"plungerpull"
	If keycode = RightFlipperKey Then Controller.Switch(7) = 1
End Sub

Sub Table1_KeyUp(ByVal KeyCode)
	If KeyUpHandler(keycode) Then Exit Sub
	If keycode = PlungerKey Then Plunger.Fire:PlaySound"plunger"
	If keycode = RightFlipperKey Then Controller.Switch(7) = 0
End Sub

'**********************************************************************************************************

'Drain hole
Sub Drain_Hit:playsound"drain":bsTrough.addball me:End Sub
Sub sw4_Hit() : bsTopEject.Addball 0 : playsound SoundFX("popper_ball",DOFContactors): End Sub
Sub sw5_Hit() : bsRightEject.Addball 0 : playsound SoundFX("popper_ball",DOFContactors): End Sub

'Stand Up Target
Sub sw17_Hit : vpmTimer.PulseSw(17) : End Sub

'Spinners
Sub sw18_Spin:vpmTimer.PulseSw 18 : playsound"fx_spinner" : End Sub

'Scoring rubber
Sub sw19_Hit(): vpmtimer.pulsesw 19 : playsound"rubber_hit_3" : End Sub
Sub sw19a_Hit(): vpmtimer.pulsesw 19 : playsound"rubber_hit_3" : End Sub

'Wire Triggers
Sub sw12_Hit:Controller.Switch(12)=1 : playsound"rollover" : End Sub 
Sub sw12_unHit:Controller.Switch(12)=0:End Sub
Sub sw13_Hit:Controller.Switch(13)=1 : playsound"rollover" : End Sub 
Sub sw13_unHit:Controller.Switch(13)=0:End Sub
Sub sw14_Hit:Controller.Switch(14)=1 : playsound"rollover" : End Sub 
Sub sw14_unHit:Controller.Switch(14)=0:End Sub

Sub sw21_Hit:Controller.Switch(21)=1 : playsound"rollover" : End Sub 
Sub sw21_unHit:Controller.Switch(21)=0:End Sub
Sub sw22_Hit:Controller.Switch(22)=1 : playsound"rollover" : End Sub 
Sub sw22_unHit:Controller.Switch(22)=0:End Sub
Sub sw23_Hit:Controller.Switch(23)=1 : playsound"rollover" : End Sub 
Sub sw23_unHit:Controller.Switch(23)=0:End Sub
Sub sw24_Hit:Controller.Switch(24)=1 : playsound"rollover" : End Sub 
Sub sw24_unHit:Controller.Switch(24)=0:End Sub

'Star Triggers
Sub sw20a_Hit:Controller.Switch(20)=1 : playsound"rollover" : End Sub 
Sub sw20a_unHit:Controller.Switch(20)=0:End Sub
Sub sw20b_Hit:Controller.Switch(20)=1 : playsound"rollover" : End Sub 
Sub sw20b_unHit:Controller.Switch(20)=0:End Sub
Sub sw20c_Hit:Controller.Switch(20)=1 : playsound"rollover" : End Sub 
Sub sw20c_unHit:Controller.Switch(20)=0:End Sub
Sub sw25_Hit:Controller.Switch(25)=1 : playsound"rollover" : End Sub 
Sub sw25_unHit:Controller.Switch(25)=0:End Sub
Sub sw26_Hit:Controller.Switch(26)=1 : playsound"rollover" : End Sub 
Sub sw26_unHit:Controller.Switch(26)=0:End Sub

'Bumpers
Sub Bumper1_Hit : vpmTimer.PulseSw(40) : playsound SoundFX("fx_bumper1",DOFContactors): End Sub
Sub Bumper2_Hit : vpmTimer.PulseSw(38) : playsound SoundFX("fx_bumper1",DOFContactors): End Sub
Sub Bumper3_Hit : vpmTimer.PulseSw(39) : playsound SoundFX("fx_bumper1",DOFContactors): End Sub

'Drop Targets
 Sub Sw27_Dropped:dtL.Hit 1 :End Sub  
 Sub Sw28_Dropped:dtL.Hit 2 :End Sub  
 Sub Sw29_Dropped:dtL.Hit 3 :End Sub
 Sub Sw30_Dropped:dtL.Hit 4 :End Sub
 Sub Sw31_Dropped:dtL.Hit 5 :End Sub
 Sub Sw32_Dropped:dtL.Hit 6 :End Sub

 Sub Sw33_Dropped:dtM.Hit 1 :End Sub
 Sub Sw34_Dropped:dtM.Hit 2 :End Sub
 Sub Sw35_Dropped:dtM.Hit 3 :End Sub

 Sub Sw42_Dropped:dtT.Hit 3 :End Sub
 Sub Sw43_Dropped:dtT.Hit 2 :End Sub
 Sub Sw44_Dropped:dtT.Hit 1 :End Sub

 Sub Sw46_Dropped:dtR.Hit 3 :End Sub
 Sub Sw47_Dropped:dtR.Hit 2 :End Sub
 Sub Sw48_Dropped:dtR.Hit 1 :End Sub

'***************************************************
'       JP's VP10 Fading Lamps & Flashers
'       Based on PD's Fading Light System
' SetLamp 0 is Off
' SetLamp 1 is On
' fading for non opacity objects is 4 steps
'***************************************************

Dim LampState(200), FadingLevel(200)
Dim FlashSpeedUp(200), FlashSpeedDown(200), FlashMin(200), FlashMax(200), FlashLevel(200)

InitLamps()             ' turn off the lights and flashers and reset them to the default parameters
LampTimer.Interval = 5 'lamp fading speed
LampTimer.Enabled = 1

' Lamp & Flasher Timers

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

'****************************************************************
'*   JPJ's inserts - reflection Mod - Lights on layer 5 - __z   *
'*   Ball parameters changed : Refl 1 to 7 - Bulb 1 to 4        *
'****************************************************************

 Sub UpdateLamps
  	NFadeLm 1, l1b
	NFadeLm 1, l1z
  	NFadeL 1, l1
  	NFadeLm 2, l2z
  	NFadeL 2, l2
  	NFadeLm 3, l3z
  	NFadeL 3, l3
  	NFadeLm 4, l4z
  	NFadeL 4, l4
  	NFadeLm 5, l5z
  	NFadeL 5, l5
  	NFadeLm 6, l6z
  	NFadeL 6, l6
  	NFadeLm 7, l7z
  	NFadeL 7, l7
 	NFadeLm 8, l8z
  	NFadeL 8, l8
  	NFadeLm 9, l9z
  	NFadeL 9, l9
  	NFadeLm 10, l10z
  	NFadeL 10, l10
  	NFadeLm 11, l11z
  	NFadeL 11, l11
' 	NFadeLm 12, L12	'Bumper3
	NFadeLm 12, Light12a
	NFadeL 12, Light12
  	'NFadeL 13, l13 'Ball In Play BG
  	NFadeL 13, GIBottomMid4 'Ball In Play BG
  	NFadeLm 14, l14z
  	NFadeL 14, l14
  	NFadeLm 15, l15z
  	NFadeL 15, l15
  	NFadeLm 17, l17z
  	NFadeL 17, l17
  	NFadeLm 18, l18z
  	NFadeL 18, l18
  	NFadeLm 19, l19z
  	NFadeL 19, l19
  	NFadeLm 20, l20z
  	NFadeL 20, l20
    NFadeLm 21, l21z
	NFadeL 21, l21
  	NFadeLm 22, l22z
  	NFadeL 22, l22
  	NFadeLm 23, l23z
  	NFadeL 23, l23
  	NFadeLm 24, l24z
  	NFadeL 24, l24
  	NFadeLm 25, l25z
  	NFadeL 25, l25
  	NFadeLm 26, l26z
  	NFadeL 26, l26
  	'NFadeL 27, l27  'Match BG
 	'NFadeLm 28, L28	'Bumper2
 	NFadeLm 28, Light28a
 	NFadeL 28, Light28
  	'NNFadeL 29, l29  'HSTD BG
   	NFadeLm 30, l30z
 	NFadeL 30, l30
  	NFadeLm 31, l31z
  	NFadeL 31, l31
  	NFadeLm 33, l33z
  	NFadeL 33, l33
   	NFadeLm 34, l34z
 	NFadeL 34, l34
    NFadeLm 35, l35z
	NFadeL 35, l35
  	NFadeLm 36, l36z
  	NFadeLm 36, l36 
  	NFadeLm 36, l36az
  	NFadeL 36, l36a
  	NFadeLm 37, l37z
  	NFadeL 37, l37
  	NFadeLm 38, l38z
  	NFadeL 38, l38
  	NFadeLm 39, l39z
  	NFadeL 39, l39
    NFadeLm 40, l40z
	NFadeL 40, l40
  	NFadeLm 41, l41z
  	NFadeL 41, l41
  	NFadeLm 42, l42z
  	NFadeL 42, l42
  	'NFadeL 43, l43  'Shoot Again BG
 	'NFadeLm 44, L44	'Bumper1
 	NFadeLm 44, Light44a
 	NFadeL 44, Light44
  	'NFadeL 45, l45  'Game Over BG
  	NFadeLm 46, l46z
  	NFadeL 46, l46
   	NFadeLm 49, l49z
 	NFadeLm 49, l49
  	NFadeLm 49, l49b
   	NFadeLm 49, l49az
   	NFadeL 49, l49a
  	NFadeLm 50, l50z
  	NFadeL 50, l50
  	NFadeLm 51, l51z
  	NFadeL 51, l51
  	NFadeLm 52, l52z
  	NFadeL 52, l52
   	NFadeLm 53, l53z
 	NFadeL 53, l53
   	NFadeLm 54, l54z
 	NFadeL 54, l54
  	NFadeLm 55, l55z
  	NFadeL 55, l55
  	NFadeLm 56, l56z
  	NFadeL 56, l56
  	NFadeLm 57, l57z
  	NFadeL 57, l57
  	NFadeLm 58, l58z
  	NFadeL 58, l58
  	'NFadeLm 59, GIBottomMid4  	
	NFadeLm 59, GIBottomMid4Upper
  	NFadeL 59, l59	'Apron Credit Light
   	NFadeLm 60, l60z
 	NFadeL 60, l60
  	'NFadeL 61, l61  'Tilt BG
   	NFadeLm 62, l62z
 	NFadeL 62, l62
   	NFadeLm 63, l63z
 	NFadeL 63, l63
  	NFadeLm 65, l65z
  	NFadeL 65, l65
  	NFadeLm 66, Light66g
  	NFadeLm 66, Light66f
  	NFadeLm 66, Light66e
  	NFadeLm 66, Light66d
  	NFadeLm 66, Light66c
	NFadeLm 66, Light66b
 	NFadeLm 66, Light66a ' JPJ It was NFadeL 66, Light66a 
  	NFadeLm 66, Light66
   	NFadeLm 66, l66z
   	NFadeL 66, l66 ' JPJ I add this line there was no assignement for this light
	NFadeLm 81, Light81a
  	NFadeLm 81, l81z
  	NFadeL 81, l81
  	NFadeLm 82, Light82g
  	NFadeLm 82, Light82d
  	NFadeLm 82, Light82c
  	NFadeLm 82, Light82b
  	NFadeLm 82, Light82
  	NFadeL 82, Light82a
	NFadeLm 97, Light97b
  	NFadeLm 97, l97z
  	NFadeL 97, l97
  	NFadeLm 98, Light98g
  	NFadeLm 98, Light98d
  	NFadeLm 98, Light98c
  	NFadeLm 98, Light98b
  	NFadeLm 98, Light98
  	NFadeL 98, Light98a
  	NFadeLm 113, Light113b
  	NFadeLm 113, l113z
  	NFadeL 113, l113

'  	NFadeL 66, Light15
'GITL001.State = 2

'****** End of JPJ's inserts ******

UpdateShadows


  End Sub

'****** Start of dynamic Light Shadows


Sub UpdateShadows

If sw35.IsDropped Then
	If sw34.IsDropped Then
		If sw33.IsDropped Then 
			LTL001.State = 1
			LTL101.State = 0
			LTL011.State = 0
			LTL111.State = 0
			
			LTR100.State = 1
			LTR101.State = 0
			LTR110.State = 0
			LTR111.State = 0
		Else
			LTL001.State = 1
			LTL101.State = 0
			LTL011.State = 0
			LTL111.State = 0		
			
			LTR100.State = 0
			LTR101.State = 1
			LTR110.State = 0
			LTR111.State = 0
		End If
	Else	If sw33.IsDropped Then
			LTL001.State = 0
			LTL101.State = 0
			LTL011.State = 1
			LTL111.State = 0
			
			LTR100.State = 0
			LTR101.State = 0
			LTR110.State = 1
			LTR111.State = 0
		Else
			LTL001.State = 0
			LTL101.State = 0
			LTL011.State = 1
			LTL111.State = 0
			
			LTR100.State = 0
			LTR101.State = 0
			LTR110.State = 0
			LTR111.State = 1
		End If
	End If

Else
	If sw34.IsDropped Then
		If sw33.IsDropped Then 
			LTL001.State = 0
			LTL101.State = 1
			LTL011.State = 0
			LTL111.State = 0
			
			LTR100.State = 1
			LTR101.State = 0
			LTR110.State = 0
			LTR111.State = 0
		Else
			LTL001.State = 0
			LTL101.State = 1
			LTL011.State = 0
			LTL111.State = 0		
			
			LTR100.State = 0
			LTR101.State = 1
			LTR110.State = 0
			LTR111.State = 0
		End If
	Else	If sw33.IsDropped Then
			LTL001.State = 0
			LTL101.State = 0
			LTL011.State = 0 
			LTL111.State = 1
			
			LTR100.State = 0
			LTR101.State = 0
			LTR110.State = 1
			LTR111.State = 0
		Else
			LTL001.State = 0
			LTL101.State = 0
			LTL011.State = 0
			LTL111.State = 1
			
			LTR100.State = 0
			LTR101.State = 0
			LTR110.State = 0
			LTR111.State = 1
		End If
	End If
End If

End Sub
'******** End of dynamic Light Shadows




' div lamp subs

Sub InitLamps()
    Dim x
    For x = 0 to 200
        LampState(x) = 0        ' current light state, independent of the fading level. 0 is off and 1 is on
        FadingLevel(x) = 4      ' used to track the fading state
        FlashSpeedUp(x) = 0.4   ' faster speed when turning on the flasher
        FlashSpeedDown(x) = 0.2 ' slower speed when turning off the flasher
        FlashMax(x) = 1         ' the maximum value when on, usually 1
        FlashMin(x) = 0         ' the minimum value when off, usually 0
        FlashLevel(x) = 0       ' the intensity of the flashers, usually from 0 to 1
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
        Case 6, 7, 8:FadingLevel(nr) = FadingLevel(nr) + 1             'wait
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


'**********************************************************************************************************
'Digital Display
'**********************************************************************************************************
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
			end if
		next
		end if
end if
End Sub


'**********************************************************************************************************
'**********************************************************************************************************



'Bally Fathom
'added by Inkochnito
Sub editDips
	Dim vpmDips : Set vpmDips = New cvpmDips
	With vpmDips
		.AddForm 700,400,"Fathom - DIP switches"
		.AddChk 7,10,120,Array("Match feature",&H08000000)'dip 28
		.AddChk 130,10,120,Array("Game over attract",&H20000000)'dip 30
		.AddChk 260,10,120,Array("Credits display",&H04000000)'dip 27
		.AddFrame 2,30,190,"Maximum credits",&H03000000,Array("10 credits",0,"15 credits",&H01000000,"25 credits",&H02000000,"40 credits",&H03000000)'dip 25&26
		.AddFrame 2,106,190,"Locked ball adjustment",&H00000020,Array("eject ball at game end",0,"ball remains locked",&H00000020)'dip 6
		.AddFrame 2,152,190,"Bonus special lit at maximum",&H00000040,Array("blue and green bonus",0,"blue or green bonus",&H00000040)'dip 7
		.AddFrame 2,198,190,"Extra ball lite is lit for",&H00000080,Array("6 seconds",0,"10 seconds",&H00000080)'dip 8
		.AddFrame 2,244,190,"A-B-C special adjust",32768,Array("replay",0,"alternating points or replay",32768)'dip 16
		.AddFrame 205,30,190,"Balls per game",&HC0000000,Array("2 balls",&HC0000000,"3 balls",0,"4 balls",&H80000000,"5 balls",&H40000000)'dip 31&32
		.AddFrame 205,106,190,"Number of replays per game",&H10000000,Array("Only 1 replay per player per game",0,"All replays earned will be collected",&H10000000)'dip 29
		.AddFrame 205,152,190,"Drop target memory",&H00200000,Array("not in memory",0,"kept in memory",&H00200000)'dip 22
		.AddFrame 205,198,190,"55,000 bonus lites",&H00400000,Array("not in memory",0,"kept in memory",&H00400000)'dip 23
		.AddFrame 205,244,190,"A-B-C lites",&H00800000,Array("not in memory",0,"kept in memory",&H00800000)'dip 24
		.AddLabel 50,300,350,20,"Set selftest position 16,17,18 and 19 to 03 for the best gameplay."
		.AddLabel 50,320,300,20,"After hitting OK, press F3 to reset game with new settings."
		.ViewDips
	End With
End Sub
Set vpmShowDips = GetRef("editDips")


' *********************************************************************
' *********************************************************************

					'Start of VPX call back Functions

' *********************************************************************
' *********************************************************************


'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep

Sub RightSlingShot_Slingshot
	vpmTimer.PulseSw 36
    PlaySound SoundFX("right_slingshot",DOFContactors), 0, 1, 0.05, 0.05
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.TransZ = -20
    RStep = 0
    RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.TransZ = -10
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.TransZ = 0:RightSlingShot.TimerEnabled = 0:
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
	vpmTimer.PulseSw 37
    PlaySound SoundFX("left_slingshot",DOFContactors),0,1,-0.05,0.05
    LSling.Visible = 0
    LSling1.Visible = 1
    sling2.TransZ = -20
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.TransZ = -10
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.TransZ = 0:LeftSlingShot.TimerEnabled = 0:
    End Select
    LStep = LStep + 1
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
	PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 2000, Pan(ball1), 0, Pitch(ball1), 0, 0
End Sub

'*****************************************
'			FLIPPER SHADOWS
'*****************************************

sub FlipperTimer_Timer()
	FlipperLSh.RotZ = LeftFlipper.currentangle
	FlipperRSh.RotZ = RightFlipper.currentangle

End Sub

'*****************************************
'			BALL SHADOW
'*****************************************
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
    ' render the shadow for each ball
    For b = 0 to UBound(BOT)
        If BOT(b).X < Table1.Width/2 Then
            BallShadow(b).X = ((BOT(b).X) - (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/7)) + 10
        Else
            BallShadow(b).X = ((BOT(b).X) + (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/7)) - 10
        End If
        ballShadow(b).Y = BOT(b).Y + 20
        If BOT(b).Z > 20 Then
            BallShadow(b).visible = 1
        Else
            BallShadow(b).visible = 0
        End If
    Next
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

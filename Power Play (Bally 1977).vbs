Option Explicit
Randomize
 
On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0
 
Const cGameName="pwerplay",UseSolenoids=2,UseLamps=0,UseGI=0,SSolenoidOn="SolOn",SSolenoidOff="SolOff", SCoin="coin"
                '"pwerplab" for freeplay
 
Dim enableBallControl
 
enableBallControl = 1   ' 1 to enable, 0 to disable
 
LoadVPM "01500100", "BALLY.VBS", 3.10
Dim DesktopMode: DesktopMode = Table1.ShowDT
 
Dim BallShadows: Ballshadows=1          '******************set to 1 to turn on Ball shadows
Dim FlipperShadows: FlipperShadows=1  '***********set to 1 to turn on Flipper shadows
 
If DesktopMode = True Then 'Show Desktop components
	Ramp16.visible=1
	Ramp15.visible=1
	Else
	Ramp16.visible=0
	Ramp15.visible=0
End if
 
'*************************************************************
'Solenoid Call backs
'**********************************************************************************************************
 
SolCallback(1)="SolPostDown"
SolCallback(2)="vpmSolSound SoundFX(""chime1"",DOFChimes),"
SolCallback(3)="vpmSolSound SoundFX(""chime2"",DOFChimes),"
SolCallback(4)="vpmSolSound SoundFX(""chime3"",DOFChimes),"
SolCallBack(5)="vpmSolSound SoundFX(""chime4"",DOFChimes),"
SolCallback(6)="vpmSolSound SoundFX(""Knocker"",DOFKnocker),"
SolCallback(7)="bsTrough.SolOut"
SolCallBack(8)="solsaucer"
SolCallback(15)="SolRDropUp"
SolCallback(13)="SolLDropUp"
SolCallback(17)="SolPostUp"
SolCallback(19)="vpmNudge.SolGameOn"
 
 
 SolCallback(sLRFlipper) = "SolRFlipper"
 SolCallback(sLLFlipper) = "SolLFlipper"
 
 
Sub SolLFlipper(Enabled)
     If Enabled Then
         PlaySound SoundFX("fx_Flipperup",DOFFlippers):LeftFlipper.RotateToEnd:LeftFlipper1.RotateToEnd
     Else
         PlaySound SoundFX("fx_Flipperdown",DOFFlippers):LeftFlipper.RotateToStart:LeftFlipper1.RotateToStart
     End If
  End Sub
 
Sub SolRFlipper(Enabled)
     If Enabled Then
         PlaySound SoundFX("fx_Flipperup",DOFFlippers):RightFlipper.RotateToEnd:RightFlipper1.RotateToEnd
     Else
         PlaySound SoundFX("fx_Flipperdown",DOFFlippers):RightFlipper.RotateToStart:RightFlipper1.RotateToStart
     End If
End Sub
 
'**********************************************************************************************************
 
'Solenoid Controlled toys
'**********************************************************************************************************
 
 Sub SolSaucer(Enable)
    If Enable then
            bsSaucer2.ExitSol_On
            RKick = 0
            kickarmtop_prim.ObjRotX = -12
            RKickTimer.Enabled = 1
    End If
 End Sub
 
Sub SolPostUp(Enabled)
    If CenterPost_prim.TransZ=0 Then PlaySound SoundFX("centerpost",DOFContactors) End If
    If Enabled Then
        SetLamp 150, 1
        Post.IsDropped=0
        flipperleft_prim.image= "leftflipperON"
        flipperright_prim.image= "rightflipperON"
        centerpost_prim.image="pp_popupON"
        CenterPost_prim.TransZ = 24
    End If
End Sub
 
Sub SolPostDown(Enabled)
    If CenterPost_prim.TransZ=24 Then PlaySound SoundFX("centerpost",DOFContactors) End If
    If Enabled Then
        SetLamp 150, 0
        Post.IsDropped=1
        flipperleft_prim.image= "leftflipperOFF"
        flipperright_prim.image= "rightflipperOFF"
        centerpost_prim.image="pp_popupOFF"
        CenterPost_prim.Transz = 0
    End If
End Sub
 
    if ballshadows=1 then
        BallShadowUpdate.enabled=1
      else
        BallShadowUpdate.enabled=0
    end if
 
    if flippershadows=1 then
        FlipperLSh.visible=1
        FlipperRSh.visible=1
        FlipperLSh1.visible=1
        FlipperRSh1.visible=1
      else
        FlipperLSh.visible=0
        FlipperRSh.visible=0
        FlipperLSh1.visible=0
        FlipperRSh1.visible=0
    end if
 
 
'Primitive Flipper Code
Sub FlipperTimer_Timer
    LUFlogo.roty = LeftFlipper.currentangle - -128
    RUFlogo.roty = RightFlipper.currentangle - -52
    flipperleft_prim.rotz = leftflipper.currentangle
    flipperright_prim.rotz = rightflipper.currentangle
    if FlipperShadows=1 then
        FlipperLsh.rotz= LeftFlipper.currentangle
        FlipperLsh1.rotz= LeftFlipper1.currentangle
        FlipperRsh.rotz= RightFlipper.currentangle
        FlipperRsh1.rotz= RightFlipper1.currentangle
    end if
	if l4.state=1 Then
	    metalguide_prim.image="pp_metalalllit"
		If l12.state=0 Then
			metalguide_prim.image="pp_metal_bottomlit"
		End If
	Else
	metalguide_prim.image="pp_metal"
		end if


	if l58.state=1 Then
		bumper1cap.image="bumper1ON"
		Else
		bumper1cap.image="bumper1OFF"
	end If
	if l42.state=1 Then
		bumper2cap.image="bumper2ON"
		bumper3cap.image="bumper3ON"
		Else
		bumper2cap.image="bumper2OFF"
		bumper3cap.image="bumper3OFF"
	end If
End Sub
 
'*****************************************
'           BALL SHADOW
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
            BallShadow(b).X = ((BOT(b).X) - (Ballsize/16) + ((BOT(b).X - (Table1.Width/2))/17))' + 13
        Else
            BallShadow(b).X = ((BOT(b).X) + (Ballsize/16) + ((BOT(b).X - (Table1.Width/2))/17))' - 13
        End If
        ballShadow(b).Y = BOT(b).Y + 10
        If BOT(b).Z > 20 Then
            BallShadow(b).visible = 1
        Else
            BallShadow(b).visible = 0
        End If
    Next
End Sub
 
 
'**********************************************************************************************************
 
'Initiate Table
'**********************************************************************************************************
 
 Dim bsTrough, bsSaucer2, dtbank1, dtbank2
 Sub Table1_Init
    vpmInit Me
    On Error Resume Next
        With Controller
        .GameName = cGameName
        If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
        .SplashInfoLine = "Power Play (Bally 1977)"&chr(13)&"v. 1.0"
        .HandleMechanics=0
        .HandleKeyboard=0
        .ShowDMDOnly=1
        .ShowFrame=0
        .ShowTitle=0
        .hidden = 1
         On Error Resume Next
         .Run GetPlayerHWnd
         If Err Then MsgBox Err.Description
         On Error Goto 0
     End With
     On Error Goto 0
 
     PinMAMETimer.Interval = PinMAMEInterval
     PinMAMETimer.Enabled = 1
 
     vpmNudge.TiltSwitch = swTilt
     vpmNudge.Sensitivity = 1
     vpmNudge.TiltObj = Array(Bumper1,Bumper2,Bumper3,LeftSlingshot,RightSlingshot)
 
     Set bsTrough = New cvpmBallStack
         bsTrough.InitSw 0, 8, 0, 0, 0, 0, 0, 0
         bsTrough.InitKick BallRelease, 80, 6
         bsTrough.InitExitSnd SoundFX("ballrelease",DOFContactors), SoundFX("Solenoid",DOFContactors)
         bsTrough.Balls = 1
 
     Set bsSaucer2 = New cvpmBallStack
         bsSaucer2.InitSaucer kicker1, 32, 150, 8
         bsSaucer2.KickAngleVar = 3
         bsSaucer2.KickForceVar = 3
         bsSaucer2.InitExitSnd SoundFX("Popper",DOFContactors), SoundFX("Solenoid",DOFContactors)
 
     set dtbank1 = new cvpmdroptarget
         dtbank1.initdrop array(sw21, sw22, sw23, sw24), array(21, 22, 23, 24)
         dtbank1.initsnd  SoundFX("DTDrop",DOFDropTargets),SoundFX("DTReset",DOFContactors)
 
     set dtbank2 = new cvpmdroptarget
         dtbank2.initdrop array(sw20, sw19, sw18, sw17), array(20, 19, 18, 17)
         dtbank2.initsnd  SoundFX("DTDrop",DOFDropTargets),SoundFX("DTReset",DOFContactors)
 
         Post.IsDropped=1
 
'*****GI Lights On
dim xx
For each xx in GI:xx.State = 1: Next
 
''*****Drop Lights Off
For each xx in DTLeftLights: xx.state=0:Next
For each xx in DTRightLights: xx.state=0:Next
 
End Sub
 
'**********************************************************************************************************
'Plunger code
'**********************************************************************************************************
 
Sub Table1_KeyDown(ByVal KeyCode)
    If KeyDownHandler(keycode) Then Exit Sub
    If keycode = PlungerKey Then Plunger.Pullback:playsound"plungerpull"
 
'************************   Start Ball Control 1/3
    if enableBallControl then
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
    End If
    if keycode = 203 then bcleft = 1        ' Left Arrow
    if keycode = 200 then bcup = 1          ' Up Arrow
    if keycode = 208 then bcdown = 1        ' Down Arrow
    if keycode = 205 then bcright = 1       ' Right Arrow
   
'************************   End Ball Control 1/3
 
End Sub
 
Sub Table1_KeyUp(ByVal KeyCode)
    If KeyUpHandler(keycode) Then Exit Sub
    If keycode = PlungerKey Then Plunger.Fire:PlaySound"plunger"
 
'************************   Start Ball Control 2/3
    if keycode = 203 then bcleft = 0        ' Left Arrow
    if keycode = 200 then bcup = 0          ' Up Arrow
    if keycode = 208 then bcdown = 0        ' Down Arrow
    if keycode = 205 then bcright = 0       ' Right Arrow
'************************   End Ball Control 2/3
 
End Sub
 
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
 
 
'**********************************************************************************************************
 
 ' Drain hole and kickers
Sub Drain_Hit:bsTrough.addball me : playsound"drain" : DOF 103, DOFPulse : End Sub
Sub kicker1_Hit:playsound"fx_balldrop":bsSaucer2.AddBall 0:End Sub
 
 ' Droptargets
 Sub sw20_Dropped:dtbank2.Hit 1
     R1DA.state=1
End Sub
 
 Sub sw19_Dropped:dtbank2.Hit 2
     R1DB.state=1
End Sub
 
 Sub sw18_Dropped:dtbank2.Hit 3
     R1DC.state=1:L1DC.state=1
End Sub
 
 Sub sw17_Dropped:dtbank2.Hit 4
End Sub
 
 Sub sw21_Dropped:dtbank1.Hit 1
End Sub
 
 Sub sw22_Dropped:dtbank1.Hit 2
    L1DC.state=1
End Sub
 
 Sub sw23_Dropped:dtbank1.Hit 3
    L1DB.state=1
End Sub
 
 Sub sw24_Dropped:dtbank1.Hit 4
    L1DA.state=1
End Sub
 
Sub SolLDropUp(enabled)
    dim xx
    if enabled then
        dtBank1.SolDropUp enabled
        For each xx in DTLeftLights: xx.state=0:Next
    end if
End Sub
 
Sub SolRDropUp(enabled)
    dim xx
    if enabled then
        dtBank2.SolDropUp enabled
        For each xx in DTRightLights: xx.state=0:Next
    end if
End Sub
 
'Rollovers
 Sub sw25a_Hit:Controller.Switch(25) = 1 : playsound"rollover" : End Sub
 Sub sw25a_UnHit:Controller.Switch(25) = 0:End Sub
 Sub sw25b_Hit:Controller.Switch(25) = 1 : playsound"rollover" : End Sub
 Sub sw25b_UnHit:Controller.Switch(25) = 0:End Sub
 Sub sw26_Hit:Controller.Switch(26) = 1 : playsound"rollover" : End Sub
 Sub sw26_UnHit:Controller.Switch(26) = 0:End Sub
 Sub sw28_Hit:Controller.Switch(28) = 1 : playsound"rollover" : End Sub
 Sub sw28_UnHit:Controller.Switch(28) = 0:End Sub
 Sub sw5_Hit:Controller.Switch(5) = 1 : playsound"rollover" : End Sub
 Sub sw5_UnHit:Controller.Switch(5) = 0:End Sub
 Sub sw35_Hit:Controller.Switch(35) = 1 : playsound"rollover" : End Sub
 Sub sw35_UnHit:Controller.Switch(35) = 0:End Sub
 Sub sw30_Hit:Controller.Switch(30) = 1 : playsound"rollover" : End Sub
 Sub sw30_UnHit:Controller.Switch(30) = 0:End Sub
 Sub sw4_Hit:Controller.Switch(4) = 1 : playsound"rollover" : End Sub
 Sub sw4_UnHit:Controller.Switch(4) = 0:End Sub
 Sub sw34_Hit:Controller.Switch(34) = 1 : playsound"rollover" : End Sub
 Sub sw34_UnHit:Controller.Switch(34) = 0:End Sub
 Sub sw29_Hit:Controller.Switch(29) = 1 : playsound"rollover" : End Sub
 Sub sw29_UnHit:Controller.Switch(29) = 0:End Sub
 Sub sw32b_Hit:Controller.Switch(33) = 1 : playsound"rollover" : End Sub
 Sub sw32b_UnHit:Controller.Switch(33) = 0:End Sub
 Sub sw27_Hit:Controller.Switch(27) = 1 : playsound"rollover" : End Sub
 Sub sw27_UnHit:Controller.Switch(27) = 0:End Sub
 Sub sw28_Hit:Controller.Switch(28) = 1 : playsound"rollover" : End Sub
 Sub sw28_UnHit:Controller.Switch(28) = 0:End Sub
 Sub sw1a_Hit:Controller.Switch(1) = 1 : playsound"rollover" : DOF 101, DOFOn : End Sub
 Sub sw1a_UnHit:Controller.Switch(1) = 0: DOF 101, DOFOff : End Sub
 Sub sw1b_Hit:Controller.Switch(1) = 1 : playsound"rollover" : DOF 102, DOFOn : End Sub
 Sub sw1b_UnHit:Controller.Switch(1) = 0: DOF 102, DOFOff : End Sub
 
'Bumpers
Sub Bumper1_Hit : vpmTimer.PulseSw(38) : playsound SoundFX("fx_bumper1",DOFContactors): End Sub
Sub Bumper2_Hit : vpmTimer.PulseSw(40) : playsound SoundFX("fx_bumper2",DOFContactors): End Sub
Sub Bumper3_Hit : vpmTimer.PulseSw(39) : playsound SoundFX("fx_bumper3",DOFContactors): End Sub
 
'Stand Up Targets
Sub sw32_Hit:vpmTimer.PulseSw (33):playsound SoundFX("target",DOFTargets):End Sub
Sub sw2_Hit:vpmTimer.PulseSw (2):playsound SoundFX("target",DOFTargets):End Sub
 
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
 
 Sub UpdateLamps()
     NFadeLm 1, l1
     NFadeL 1, l1a
     NFadeLm 2, l2
     NFadeL 2, l2a
     NFadeLm 3, l3
     NFadeL 3, l3a
     NFadeLm 4, l4
     NFadeL 4, l4a
     NFadeLm 5, l5
     NFadeL 5, l5a
     NFadeLm 6, l6
     NFadeL 6, l6a
     NFadeLm 7, l7
     NFadeL 7, l7a
     NFadeLm 8, l8
     NFadeL 8, l8a
     NFadeLm 11, l11
     NFadeL 11, l11a
     NFadeLm 12, l12
     NFadeLm 12, l12a
     NFadeLm 12, l12c
     NFadeLm 12, l12d
     NFadeLm 17, l17
     NFadeL 17, l17a
     NFadeLm 18, l18
     NFadeL 18, l18a
     NFadeLm 19, l19
     NFadeL 19, l19a
     NFadeLm 20, l20
     NFadeL 20, l20a
     NFadeLm 21, l21
     NFadeL 21, l21a
     NFadeLm 23, l23
     NFadeL 23, l23a
     NFadeLm 25, l25
     NFadeL 25, l25a
     NFadeLm 28, l28
     NFadeLm 28, l28a
     NFadeLm 28, l28c
     NFadeLm 28, l28d
     NFadeLm 33, l33
     NFadeL 33, l33a
     NFadeLm 34, l34
     NFadeL 34, l34a
     NFadeLm 35, l35
     NFadeL 35, l35a
     NFadeLm 37, l37
     NFadeL 37, l37a
     NFadeLm 39, l39
     NFadeL 39, l39a
     NFadeLm 41, l41
     NFadeL 41, l41a
     NFadeLm 42, l42
     NFadeLm 42, l42a
	 NFadeLm 42, L42b
	 NFadeLm 42, L42c
     NFadeLm 44, l44
     NFadeL 44, l44a
     NFadeLm 49, l49
     NFadeL 49, l49a
     NFadeLm 52, l52
     NFadeL 52, l52a
     NFadeLm 53, l53
     NFadeL 53, l53a
     NFadeLm 55, l55
     NFadeL 55, l55a
     NFadeLm 57, l57
     NFadeL 57, l57a
     NFadeLm 58, l58
     NFadeL 58, l58a
     NFadeL 59, l59 'Apron credit
     NFadeLm 60, l60
     NFadeL 60, l60a
     NFadeLm 150, l150
 End Sub
 
 
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
 
 'Reels
Sub FadeReel(nr, reel)
    Select Case FadingLevel(nr)
        Case 2:FadingLevel(nr) = 0
        Case 3:FadingLevel(nr) = 2
        Case 4:reel.Visible = 0:FadingLevel(nr) = 3
        Case 5:reel.Visible = 1:FadingLevel(nr) = 1
    End Select
End Sub
 
 'Inverted Reels
Sub FadeIReel(nr, reel)
    Select Case FadingLevel(nr)
        Case 2:FadingLevel(nr) = 0
        Case 3:FadingLevel(nr) = 2
        Case 4:reel.Visible = 1:FadingLevel(nr) = 3
        Case 5:reel.Visible = 0:FadingLevel(nr) = 1
    End Select
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
 
'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep, RKick, LA, RC, LC, RA, RB, RD
 
Sub RightSlingShot_Slingshot
    vpmtimer.pulsesw 36
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
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.TransZ = 0:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub
 
Sub LeftSlingShot_Slingshot
    vpmtimer.pulsesw 37
    PlaySound SoundFX("left_slingshot",DOFContactors),0,1,-0.05,0.05
    LSling.Visible = 0
    LSling1.Visible = 1
    sling2.TransZ = -32
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
End Sub
 
Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.TransZ = -10
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.TransZ = 0:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub
 
Sub WallLA_Hit
    RubberLA.Visible = 0
    RubberLA1.Visible = 1
    LA = 0
    WallLA.TimerEnabled = 1
End Sub
 
Sub WallLA_Timer
    Select Case LA
        Case 3:RubberLA1.Visible = 0:RubberLA2.Visible = 1
        Case 4:RubberLA2.Visible = 0:RubberLA.Visible = 1:WallLA.TimerEnabled = 0:
    End Select
    LA = LA + 1
End Sub
'
Sub WallLC_Hit
    RubberLC.Visible = 0
    RubberLC1.Visible = 1
    LC = 0
    WallLC.TimerEnabled = 1
End Sub
 
Sub WallLC_Timer
    Select Case LC
        Case 3:RubberLC1.Visible = 0:RubberLC2.Visible = 1
        Case 4:RubberLC2.Visible = 0:RubberLC.Visible = 1:WallLC.TimerEnabled = 0:
    End Select
    LC = LC + 1
End Sub
'
Sub WallRA_Hit
    RubberRA.Visible = 0
    RubberRA1.Visible = 1
    RA = 0
    WallRA.TimerEnabled = 1
End Sub
 
Sub WallRA_Timer
    Select Case RA
        Case 3:RubberRA1.Visible = 0:RubberRA2.Visible = 1
        Case 4:RubberRA2.Visible = 0:RubberRA.Visible = 1:WallRA.TimerEnabled = 0:
    End Select
    RA = RA + 1
End Sub
'
Sub WallRC_Hit
    RubberRC.Visible = 0
    RubberRC1.Visible = 1
    RC = 0
    WallRC.TimerEnabled = 1
End Sub
 
Sub WallRC_Timer
    Select Case RC
        Case 3:RubberRC1.Visible = 0:RubberRC2.Visible = 1
        Case 4:RubberRC2.Visible = 0:RubberRC.Visible = 1:WallRC.TimerEnabled = 0:
    End Select
    RC = RC + 1
End Sub
'
Sub sw3a_Hit
    DOF 105, DOFPulse
    vpmTimer.PulseSw 3
    RubberDL.Visible = 0
    RubberDL1.Visible = 1
    RB = 0
    sw3a.TimerEnabled = 1
End Sub
 
Sub sw3a_Timer
    Select Case RB
        Case 3:RubberDL1.Visible = 0:RubberDL2.Visible = 1
        Case 4:RubberDL2.Visible = 0:RubberDL.Visible = 1:sw3a.TimerEnabled = 0:
    End Select
    RB = RB + 1
End Sub
'
Sub sw3b_Slingshot
    DOF 106, DOFPulse
    vpmTimer.PulseSw 3
    RubberDR.Visible = 0
    RubberDR1.Visible = 1
    RD = 0
    sw3b.TimerEnabled = 1
End Sub
 
Sub sw3b_Timer
    Select Case RD
        Case 3:RubberDR1.Visible = 0:RubberDR2.Visible = 1
        Case 4:RubberDR2.Visible = 0:RubberDR.Visible = 1:sw3b.TimerEnabled = 0:
    End Select
    RD = RD + 1
End Sub
 
 
Sub RKickTimer_Timer
    Select Case RKick
        Case 1:kickarmtop_prim.ObjRotX = -50
        Case 2:kickarmtop_prim.ObjRotX = -50
        Case 3:kickarmtop_prim.ObjRotX = -50
        Case 4:kickarmtop_prim.ObjRotX = -50
        Case 5:kickarmtop_prim.ObjRotX = -50
        Case 6:kickarmtop_prim.ObjRotX = -50
        Case 7:kickarmtop_prim.ObjRotX = -50
        Case 8:kickarmtop_prim.ObjRotX = -50
        Case 9:kickarmtop_prim.ObjRotX = -50
        Case 10:kickarmtop_prim.ObjRotX = -50
        Case 11:kickarmtop_prim.ObjRotX = -24
        Case 12:kickarmtop_prim.ObjRotX = -12
        Case 13:kickarmtop_prim.ObjRotX = 0:RKickTimer.Enabled = 0
    End Select
    RKick = RKick + 1
End Sub
 
' *********************************************************************
'                      Supporting Ball & Sound Functions
' *********************************************************************
 
Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
    Vol = Csng(BallVel(ball) ^2 / 1000)
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
 
'Bally Power Play
'added by Inkochnito
Sub editDips
    Dim vpmDips : Set vpmDips = New cvpmDips
    With vpmDips
        .AddForm 700,400,"Power Play - DIP switches"
        .AddFrame 2,20,190,"Maximum credits",&H00070000,Array("10 credits",&H00010000,"20 credits",&H00030000,"30 credits",&H00050000,"40 credits",&H00070000)'dip 17&18&19
        .AddFrame 2,99,190,"High game to date award",&H00000060,Array("no award",0,"1 credit",&H00000020,"2 credits",&H00000040,"3 credits",&H00000060)'dip 6&7
        .AddFrame 2,173,190,"Drop target bank award",&H00002000,Array("2X, 3X, 5X, extra ball, special",0,"2X, 3X, 5X && extra ball, special",&H00002000)'dip 14
        .AddFrame 2,220,190,"Knocking down 4 targets of 1 bank",&H00200000,Array("will reset both banks",0,"will reset only that bank",&H00200000)'dip 22
        .AddFrame 205,20,190,"Thumper-bumper adjust",&H00400000,Array("alternating",0,"all bumpers on",&H00400000)'dip 23
        .AddFrame 205,66,190,"High score feature",&HC0000000,Array("no award",0,"extra ball",&H80000000,"replay",&HC0000000)'dip 31&32
        .AddFrame 205,127,190,"Left && right alley rollover buttons",&H00004000,Array("scores 100 points",0,"scores 1000 points",&H00004000)'dip 15
        .AddFrame 205,173,190,"Balls per game", 32768,Array("3 balls",0,"5 balls", 32768)'dip 16
        .AddFrame 205,220,190,"Outlane special adjust",&H30000000,Array("special does NOT lite",0,"special alternates",&H20000000,"both specials on",&H30000000)'dip 29&30
        .AddChk 2,0,100,Array("Match feature",&H00100000)'dip 21
        .AddChk 150,0,100,Array("Credits display",&H00080000)'dip 20
        .AddChk 295,0,100,Array("Melody option",&H00000080)'dip 8
        .AddLabel 50,290,300,20,"After hitting OK, press F3 to reset game with new settings."
        .ViewDips
    End With
End Sub
Set vpmShowDips = GetRef("editDips")
 
Sub table1_Paused:Controller.Pause = 1:End Sub
Sub table1_unPaused:Controller.Pause = 0:End Sub
 
    Sub table1_Exit:Controller.Stop:End Sub

'Option Explicit
'Randomize
'
'On Error Resume Next
'ExecuteGlobal GetTextFile("controller.vbs")
'If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
'On Error Goto 0
'
'Const cGameName="pwerplay",UseSolenoids=2,UseLamps=0,UseGI=0,SSolenoidOn="SolOn",SSolenoidOff="SolOff", SCoin="coin"
'				'"pwerplab" for freeplay
'
'Dim enableBallControl
'
'enableBallControl = 0 	' 1 to enable, 0 to disable
'
'LoadVPM "01500100", "BALLY.VBS", 3.10
'Dim DesktopMode: DesktopMode = Table1.ShowDT
'
'Dim BallShadows: Ballshadows=1  		'******************set to 1 to turn on Ball shadows
'Dim FlipperShadows: FlipperShadows=1  '***********set to 1 to turn on Flipper shadows
'
'If DesktopMode = True Then 'Show Desktop components
'Ramp16.visible=1
'Ramp15.visible=1
'Else
'Ramp16.visible=0
'Ramp15.visible=0
'End if
'
''*************************************************************
''Solenoid Call backs
''**********************************************************************************************************
'
'SolCallback(1)="SolPostDown"
'SolCallback(2)="vpmSolSound ""chime1"","
'SolCallback(3)="vpmSolSound ""chime2"","
'SolCallback(4)="vpmSolSound ""chime3"","
'SolCallBack(5)="vpmSolSound ""chime4"","
'SolCallback(6)="vpmSolSound SoundFX(""Knocker"",DOFKnocker),"
'SolCallback(7)="bsTrough.SolOut"
'SolCallBack(8)="solsaucer"
'SolCallback(15)="SolRDropUp"
'SolCallback(13)="SolLDropUp"
'SolCallback(17)="SolPostUp"
'SolCallback(19)="vpmNudge.SolGameOn"
'
'
' SolCallback(sLRFlipper) = "SolRFlipper"
' SolCallback(sLLFlipper) = "SolLFlipper"
'
'
'Sub SolLFlipper(Enabled)
'     If Enabled Then
'         PlaySound SoundFX("fx_Flipperup",DOFContactors):LeftFlipper.RotateToEnd:LeftFlipper1.RotateToEnd
'     Else
'         PlaySound SoundFX("fx_Flipperdown",DOFContactors):LeftFlipper.RotateToStart:LeftFlipper1.RotateToStart
'     End If
'  End Sub
'  
'Sub SolRFlipper(Enabled)
'     If Enabled Then
'         PlaySound SoundFX("fx_Flipperup",DOFContactors):RightFlipper.RotateToEnd:RightFlipper1.RotateToEnd
'     Else
'         PlaySound SoundFX("fx_Flipperdown",DOFContactors):RightFlipper.RotateToStart:RightFlipper1.RotateToStart
'     End If
'End Sub
'
''**********************************************************************************************************
'
''Solenoid Controlled toys
''**********************************************************************************************************
'
' Sub SolSaucer(Enable)
'	If Enable then
'			bsSaucer2.ExitSol_On
'			RKick = 0
'			kickarmtop_prim.ObjRotX = -12
'			RKickTimer.Enabled = 1
'	End If
' End Sub
' 
'Sub SolPostUp(Enabled)
'	If CenterPost_prim.TransZ=0 Then PlaySound "centerpost" End If
'	If Enabled Then
'		SetLamp 150, 1
'		Post.IsDropped=0
'		flipperleft_prim.image= "leftflipperON"
'		flipperright_prim.image= "rightflipperON"
'		centerpost_prim.image="pp_popupON"
'		CenterPost_prim.TransZ = 24
'	End If
'End Sub
'
'Sub SolPostDown(Enabled)
'	If CenterPost_prim.TransZ=24 Then PlaySound "centerpost" End If
'	If Enabled Then
'		SetLamp 150, 0
'		Post.IsDropped=1
'		flipperleft_prim.image= "leftflipperOFF"
'		flipperright_prim.image= "rightflipperOFF"
'		centerpost_prim.image="pp_popupOFF"
'		CenterPost_prim.Transz = 0
'	End If
'End Sub
'
'	if ballshadows=1 then
'		BallShadowUpdate.enabled=1
'	  else
'		BallShadowUpdate.enabled=0
'	end if
'
'	if flippershadows=1 then 
'		FlipperLSh.visible=1
'		FlipperRSh.visible=1
'		FlipperLSh1.visible=1
'		FlipperRSh1.visible=1
'	  else
'		FlipperLSh.visible=0
'		FlipperRSh.visible=0
'		FlipperLSh1.visible=0
'		FlipperRSh1.visible=0
'	end if
'
'
''Primitive Flipper Code
'Sub FlipperTimer_Timer
'	LUFlogo.roty = LeftFlipper.currentangle - -128
'	RUFlogo.roty = RightFlipper.currentangle - -52
'	flipperleft_prim.rotz = leftflipper.currentangle
'	flipperright_prim.rotz = rightflipper.currentangle
'	if FlipperShadows=1 then
'		FlipperLsh.rotz= LeftFlipper.currentangle
'		FlipperLsh1.rotz= LeftFlipper1.currentangle - 12
'		FlipperRsh.rotz= RightFlipper.currentangle
'		FlipperRsh1.rotz= RightFlipper1.currentangle - -12
'	end if
'
'	if l12.state=1 Then
'		metalguide_prim.image="pp_metalalllit"
'		Else
'		metalguide_prim.image="pp_metal"
'	end if
'End Sub
'
''*****************************************
''			BALL SHADOW
''*****************************************
'Dim BallShadow
'BallShadow = Array (BallShadow1,BallShadow2,BallShadow3,BallShadow4,BallShadow5)
'
'Sub BallShadowUpdate_timer()
'    Dim BOT, b
'    BOT = GetBalls
'    ' hide shadow of deleted balls
'    If UBound(BOT)<(tnob-1) Then
'        For b = (UBound(BOT) + 1) to (tnob-1)
'            BallShadow(b).visible = 0
'        Next
'    End If
'    ' exit the Sub if no balls on the table
'    If UBound(BOT) = -1 Then Exit Sub
'    ' render the shadow for each ball
'    For b = 0 to UBound(BOT)
'        If BOT(b).X < Table1.Width/2 Then
'            BallShadow(b).X = ((BOT(b).X) - (Ballsize/16) + ((BOT(b).X - (Table1.Width/2))/17))' + 13
'        Else
'            BallShadow(b).X = ((BOT(b).X) + (Ballsize/16) + ((BOT(b).X - (Table1.Width/2))/17))' - 13
'        End If
'        ballShadow(b).Y = BOT(b).Y + 10
'        If BOT(b).Z > 20 Then
'            BallShadow(b).visible = 1
'        Else
'            BallShadow(b).visible = 0
'        End If
'    Next
'End Sub
'
'
''**********************************************************************************************************
'
''Initiate Table
''**********************************************************************************************************
'
' Dim bsTrough, bsSaucer2, dtbank1, dtbank2
' Sub Table1_Init
'	vpmInit Me
'	On Error Resume Next
'		With Controller
'		.GameName = cGameName
'		If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
'		.SplashInfoLine = "Power Play (Bally 1977)"&chr(13)&"v. 1.0"
'		.HandleMechanics=0
'		.HandleKeyboard=0
'		.ShowDMDOnly=1
'		.ShowFrame=0
'		.ShowTitle=0
'        .hidden = 1
'         On Error Resume Next
'         .Run GetPlayerHWnd
'         If Err Then MsgBox Err.Description
'         On Error Goto 0
'     End With
'     On Error Goto 0
' 
'     PinMAMETimer.Interval = PinMAMEInterval
'     PinMAMETimer.Enabled = 1
'
'     vpmNudge.TiltSwitch = swTilt
'     vpmNudge.Sensitivity = 1
'     vpmNudge.TiltObj = Array(Bumper1,Bumper2,Bumper3,LeftSlingshot,RightSlingshot)
' 
'     Set bsTrough = New cvpmBallStack
'         bsTrough.InitSw 0, 8, 0, 0, 0, 0, 0, 0
'         bsTrough.InitKick BallRelease, 80, 6
'         bsTrough.InitExitSnd SoundFX("ballrelease",DOFContactors), SoundFX("Solenoid",DOFContactors)
'         bsTrough.Balls = 1
' 
'     Set bsSaucer2 = New cvpmBallStack
'         bsSaucer2.InitSaucer kicker1, 32, 150, 8
'         bsSaucer2.KickAngleVar = 3
'         bsSaucer2.KickForceVar = 3
'         bsSaucer2.InitExitSnd SoundFX("Popper",DOFContactors), SoundFX("Solenoid",DOFContactors)
' 
'     set dtbank1 = new cvpmdroptarget
'         dtbank1.initdrop array(sw21, sw22, sw23, sw24), array(21, 22, 23, 24)
'         dtbank1.initsnd  SoundFX("DTDrop",DOFContactors),SoundFX("DTReset",DOFContactors)
' 
'     set dtbank2 = new cvpmdroptarget
'         dtbank2.initdrop array(sw20, sw19, sw18, sw17), array(20, 19, 18, 17)
'         dtbank2.initsnd  SoundFX("DTDrop",DOFContactors),SoundFX("DTReset",DOFContactors)
' 
'		 Post.IsDropped=1 
'
''*****GI Lights On
'dim xx
'For each xx in GI:xx.State = 1: Next
'
'''*****Drop Lights Off
'For each xx in DTLeftLights: xx.state=0:Next
'For each xx in DTRightLights: xx.state=0:Next
'
'End Sub
' 
''**********************************************************************************************************
''Plunger code
''**********************************************************************************************************
'
'Sub Table1_KeyDown(ByVal KeyCode)
'	If KeyDownHandler(keycode) Then Exit Sub
'	If keycode = PlungerKey Then Plunger.Pullback:playsound"plungerpull"
'
''************************   Start Ball Control 1/3
'	if enableBallControl then
'		if keycode = 46 then	 			' C Key
'			If contball = 1 Then
'				contball = 0
'			Else
'				contball = 1
'			End If
'		End If
'		if keycode = 48 then 				'B Key
'			If bcboost = 1 Then
'				bcboost = bcboostmulti
'			Else
'				bcboost = 1
'			End If
'		End If
'	End If
'	if keycode = 203 then bcleft = 1		' Left Arrow
'	if keycode = 200 then bcup = 1			' Up Arrow
'	if keycode = 208 then bcdown = 1		' Down Arrow
'	if keycode = 205 then bcright = 1		' Right Arrow
'	
''************************   End Ball Control 1/3
'
'End Sub
'
'Sub Table1_KeyUp(ByVal KeyCode)
'	If KeyUpHandler(keycode) Then Exit Sub
'	If keycode = PlungerKey Then Plunger.Fire:PlaySound"plunger"
'
''************************   Start Ball Control 2/3
'	if keycode = 203 then bcleft = 0		' Left Arrow
'	if keycode = 200 then bcup = 0			' Up Arrow
'	if keycode = 208 then bcdown = 0		' Down Arrow
'	if keycode = 205 then bcright = 0		' Right Arrow
''************************   End Ball Control 2/3
'
'End Sub
'
''************************   Start Ball Control 3/3
'Sub StartControl_Hit()
'	Set ControlBall = ActiveBall
'	contballinplay = true
'End Sub
'
'Sub StopControl_Hit()
'	contballinplay = false
'End Sub	
'
'Dim bcup, bcdown, bcleft, bcright, contball, contballinplay, ControlBall, bcboost
'Dim bcvel, bcyveloffset, bcboostmulti
'
'bcboost = 1		'Do Not Change - default setting
'bcvel = 4		'Controls the speed of the ball movement
'bcyveloffset = -0.01 	'Offsets the force of gravity to keep the ball from drifting vertically on the table, should be negative
'bcboostmulti = 3	'Boost multiplier to ball veloctiy (toggled with the B key) 
'
'Sub BallControl_Timer()
'	If Contball and ContBallInPlay then
'		If bcright = 1 Then
'			ControlBall.velx = bcvel*bcboost
'		ElseIf bcleft = 1 Then
'			ControlBall.velx = - bcvel*bcboost
'		Else
'			ControlBall.velx=0
'		End If
'
'		If bcup = 1 Then
'			ControlBall.vely = -bcvel*bcboost
'		ElseIf bcdown = 1 Then
'			ControlBall.vely = bcvel*bcboost
'		Else
'			ControlBall.vely= bcyveloffset
'		End If
'	End If
'End Sub
''************************   End Ball Control 3/3
'
'
''**********************************************************************************************************
' 
' ' Drain hole and kickers
'Sub Drain_Hit:bsTrough.addball me : playsound"drain" : End Sub
'Sub kicker1_Hit:playsound"fx_balldrop":bsSaucer2.AddBall 0:End Sub
'
' ' Droptargets
' Sub sw20_Dropped:dtbank2.Hit 1
'	 R1DA.state=1
'End Sub
'
' Sub sw19_Dropped:dtbank2.Hit 2
'	 R1DB.state=1
'End Sub
'
' Sub sw18_Dropped:dtbank2.Hit 3
'	 R1DC.state=1:L1DC.state=1
'End Sub
'
' Sub sw17_Dropped:dtbank2.Hit 4
'End Sub
'
' Sub sw21_Dropped:dtbank1.Hit 1
'End Sub
'
' Sub sw22_Dropped:dtbank1.Hit 2
'	L1DC.state=1
'End Sub
'
' Sub sw23_Dropped:dtbank1.Hit 3
'	L1DB.state=1
'End Sub
'
' Sub sw24_Dropped:dtbank1.Hit 4
'	L1DA.state=1
'End Sub
'
'Sub SolLDropUp(enabled)
'	dim xx
'	if enabled then
'		dtBank1.SolDropUp enabled
'		For each xx in DTLeftLights: xx.state=0:Next
'	end if
'End Sub
'
'Sub SolRDropUp(enabled)
'	dim xx
'	if enabled then
'		dtBank2.SolDropUp enabled
'		For each xx in DTRightLights: xx.state=0:Next
'	end if
'End Sub
'
''Rollovers
' Sub sw25a_Hit:Controller.Switch(25) = 1 : playsound"rollover" : End Sub 
' Sub sw25a_UnHit:Controller.Switch(25) = 0:End Sub
' Sub sw25b_Hit:Controller.Switch(25) = 1 : playsound"rollover" : End Sub 
' Sub sw25b_UnHit:Controller.Switch(25) = 0:End Sub
' Sub sw28_Hit:Controller.Switch(28) = 1 : playsound"rollover" : End Sub 
' Sub sw28_UnHit:Controller.Switch(28) = 0:End Sub
' Sub sw5_Hit:Controller.Switch(5) = 1 : playsound"rollover" : End Sub 
' Sub sw5_UnHit:Controller.Switch(5) = 0:End Sub
' Sub sw35_Hit:Controller.Switch(35) = 1 : playsound"rollover" : End Sub 
' Sub sw35_UnHit:Controller.Switch(35) = 0:End Sub
' Sub sw30_Hit:Controller.Switch(30) = 1 : playsound"rollover" : End Sub 
' Sub sw30_UnHit:Controller.Switch(30) = 0:End Sub
' Sub sw4_Hit:Controller.Switch(4) = 1 : playsound"rollover" : End Sub 
' Sub sw4_UnHit:Controller.Switch(4) = 0:End Sub
' Sub sw34_Hit:Controller.Switch(34) = 1 : playsound"rollover" : End Sub 
' Sub sw34_UnHit:Controller.Switch(34) = 0:End Sub
' Sub sw29_Hit:Controller.Switch(29) = 1 : playsound"rollover" : End Sub 
' Sub sw29_UnHit:Controller.Switch(29) = 0:End Sub
' Sub sw32b_Hit:Controller.Switch(33) = 1 : playsound"rollover" : End Sub 
' Sub sw32b_UnHit:Controller.Switch(33) = 0:End Sub
' Sub sw27_Hit:Controller.Switch(27) = 1 : playsound"rollover" : End Sub 
' Sub sw27_UnHit:Controller.Switch(27) = 0:End Sub
' Sub sw28_Hit:Controller.Switch(28) = 1 : playsound"rollover" : End Sub 
' Sub sw28_UnHit:Controller.Switch(28) = 0:End Sub
' Sub sw1a_Hit:Controller.Switch(1) = 1 : playsound"rollover" : End Sub 
' Sub sw1a_UnHit:Controller.Switch(1) = 0:End Sub
' Sub sw1b_Hit:Controller.Switch(1) = 1 : playsound"rollover" : End Sub 
' Sub sw1b_UnHit:Controller.Switch(1) = 0:End Sub
'
''Bumpers
'Sub Bumper1_Hit : vpmTimer.PulseSw(38) : playsound SoundFX("fx_bumper1",DOFContactors): End Sub
'Sub Bumper2_Hit : vpmTimer.PulseSw(40) : playsound SoundFX("fx_bumper2",DOFContactors): End Sub
'Sub Bumper3_Hit : vpmTimer.PulseSw(39) : playsound SoundFX("fx_bumper3",DOFContactors): End Sub
'
''Stand Up Targets
'Sub sw32_Hit:vpmTimer.PulseSw (33):playsound"target":End Sub
'Sub sw2_Hit:vpmTimer.PulseSw (2):playsound"target":End Sub
'
''***************************************************
''       JP's VP10 Fading Lamps & Flashers
''       Based on PD's Fading Light System
'' SetLamp 0 is Off
'' SetLamp 1 is On
'' fading for non opacity objects is 4 steps
''***************************************************
'
'Dim LampState(200), FadingLevel(200)
'Dim FlashSpeedUp(200), FlashSpeedDown(200), FlashMin(200), FlashMax(200), FlashLevel(200)
'
'InitLamps()             ' turn off the lights and flashers and reset them to the default parameters
'LampTimer.Interval = 5 'lamp fading speed
'LampTimer.Enabled = 1
'
'' Lamp & Flasher Timers
'
'Sub LampTimer_Timer()
'    Dim chgLamp, num, chg, ii
'    chgLamp = Controller.ChangedLamps
'    If Not IsEmpty(chgLamp) Then
'        For ii = 0 To UBound(chgLamp)
'            LampState(chgLamp(ii, 0) ) = chgLamp(ii, 1)       'keep the real state in an array
'            FadingLevel(chgLamp(ii, 0) ) = chgLamp(ii, 1) + 4 'actual fading step
'        Next
'    End If
'    UpdateLamps
'End Sub
'
' Sub UpdateLamps()
'     NFadeLm 1, l1
'     NFadeL 1, l1a
'     NFadeLm 2, l2
'     NFadeL 2, l2a
'     NFadeLm 3, l3
'     NFadeL 3, l3a
'     NFadeLm 4, l4
'     NFadeL 4, l4a
'     NFadeLm 5, l5
'     NFadeL 5, l5a
'     NFadeLm 6, l6
'     NFadeL 6, l6a
'     NFadeLm 7, l7
'     NFadeL 7, l7a
'     NFadeLm 8, l8
'     NFadeL 8, l8a
'     NFadeLm 11, l11
'     NFadeL 11, l11a
'     NFadeLm 12, l12
'     NFadeLm 12, l12a
'     NFadeLm 12, l12c
'     NFadeLm 12, l12d
'     NFadeLm 17, l17
'     NFadeL 17, l17a
'     NFadeLm 18, l18
'     NFadeL 18, l18a
'     NFadeLm 19, l19
'     NFadeL 19, l19a
'     NFadeLm 20, l20
'     NFadeL 20, l20a
'     NFadeLm 21, l21
'     NFadeL 21, l21a
'     NFadeLm 23, l23
'     NFadeL 23, l23a
'     NFadeLm 25, l25
'     NFadeL 25, l25a
'     NFadeLm 28, l28
'     NFadeLm 28, l28a
'     NFadeLm 28, l28c
'     NFadeLm 28, l28d
'     NFadeLm 33, l33
'     NFadeL 33, l33a
'     NFadeLm 34, l34
'     NFadeL 34, l34a
'     NFadeLm 35, l35
'     NFadeL 35, l35a
'     NFadeLm 37, l37
'     NFadeL 37, l37a
'     NFadeLm 39, l39
'     NFadeL 39, l39a
'     NFadeLm 41, l41
'     NFadeL 41, l41a
'     NFadeLm 42, l42
'     NFadeLm 42, l42a
'	 NFadeLm 42, L42b
'	 NFadeLm 42, L42c
'     NFadeLm 44, l44
'     NFadeL 44, l44a
'     NFadeLm 49, l49
'     NFadeL 49, l49a
'     NFadeLm 52, l52
'     NFadeL 52, l52a
'     NFadeLm 53, l53
'     NFadeL 53, l53a
'     NFadeLm 55, l55
'     NFadeL 55, l55a
'     NFadeLm 57, l57
'     NFadeL 57, l57a
'     NFadeLm 58, l58
'     NFadeL 58, l58a
'	 NFadeL 59, l59 'Apron credit
'     NFadeLm 60, l60
'     NFadeL 60, l60a
'	 NFadeLm 150, l150
' End Sub
'
'
'' div lamp subs
'
'Sub InitLamps()
'    Dim x
'    For x = 0 to 200
'        LampState(x) = 0        ' current light state, independent of the fading level. 0 is off and 1 is on
'        FadingLevel(x) = 4      ' used to track the fading state
'        FlashSpeedUp(x) = 0.4   ' faster speed when turning on the flasher
'        FlashSpeedDown(x) = 0.2 ' slower speed when turning off the flasher
'        FlashMax(x) = 1         ' the maximum value when on, usually 1
'        FlashMin(x) = 0         ' the minimum value when off, usually 0
'        FlashLevel(x) = 0       ' the intensity of the flashers, usually from 0 to 1
'    Next
'End Sub
'
'Sub AllLampsOff
'    Dim x
'    For x = 0 to 200
'        SetLamp x, 0
'    Next
'End Sub
'
'Sub SetLamp(nr, value)
'    If value <> LampState(nr) Then
'        LampState(nr) = abs(value)
'        FadingLevel(nr) = abs(value) + 4
'    End If
'End Sub
'
'' Lights: used for VP10 standard lights, the fading is handled by VP itself
'
'Sub NFadeL(nr, object)
'    Select Case FadingLevel(nr)
'        Case 4:object.state = 0:FadingLevel(nr) = 0
'        Case 5:object.state = 1:FadingLevel(nr) = 1
'    End Select
'End Sub
'
'Sub NFadeLm(nr, object) ' used for multiple lights
'    Select Case FadingLevel(nr)
'        Case 4:object.state = 0
'        Case 5:object.state = 1
'    End Select
'End Sub
'
''Lights, Ramps & Primitives used as 4 step fading lights
''a,b,c,d are the images used from on to off
'
'Sub FadeObj(nr, object, a, b, c, d)
'    Select Case FadingLevel(nr)
'        Case 4:object.image = b:FadingLevel(nr) = 6                   'fading to off...
'        Case 5:object.image = a:FadingLevel(nr) = 1                   'ON
'        Case 6, 7, 8:FadingLevel(nr) = FadingLevel(nr) + 1             'wait
'        Case 9:object.image = c:FadingLevel(nr) = FadingLevel(nr) + 1 'fading...
'        Case 10, 11, 12:FadingLevel(nr) = FadingLevel(nr) + 1         'wait
'        Case 13:object.image = d:FadingLevel(nr) = 0                  'Off
'    End Select
'End Sub
'
'Sub FadeObjm(nr, object, a, b, c, d)
'    Select Case FadingLevel(nr)
'        Case 4:object.image = b
'        Case 5:object.image = a
'        Case 9:object.image = c
'        Case 13:object.image = d
'    End Select
'End Sub
'
'Sub NFadeObj(nr, object, a, b)
'    Select Case FadingLevel(nr)
'        Case 4:object.image = b:FadingLevel(nr) = 0 'off
'        Case 5:object.image = a:FadingLevel(nr) = 1 'on
'    End Select
'End Sub
'
'Sub NFadeObjm(nr, object, a, b)
'    Select Case FadingLevel(nr)
'        Case 4:object.image = b
'        Case 5:object.image = a
'    End Select
'End Sub
'
'' Flasher objects
'
'Sub Flash(nr, object)
'    Select Case FadingLevel(nr)
'        Case 4 'off
'            FlashLevel(nr) = FlashLevel(nr) - FlashSpeedDown(nr)
'            If FlashLevel(nr) < FlashMin(nr) Then
'                FlashLevel(nr) = FlashMin(nr)
'                FadingLevel(nr) = 0 'completely off
'            End if
'            Object.IntensityScale = FlashLevel(nr)
'        Case 5 ' on
'            FlashLevel(nr) = FlashLevel(nr) + FlashSpeedUp(nr)
'            If FlashLevel(nr) > FlashMax(nr) Then
'                FlashLevel(nr) = FlashMax(nr)
'                FadingLevel(nr) = 1 'completely on
'            End if
'            Object.IntensityScale = FlashLevel(nr)
'    End Select
'End Sub
'
'Sub Flashm(nr, object) 'multiple flashers, it just sets the flashlevel
'    Object.IntensityScale = FlashLevel(nr)
'End Sub
'
' 'Reels
'Sub FadeReel(nr, reel)
'    Select Case FadingLevel(nr)
'        Case 2:FadingLevel(nr) = 0
'        Case 3:FadingLevel(nr) = 2
'        Case 4:reel.Visible = 0:FadingLevel(nr) = 3
'        Case 5:reel.Visible = 1:FadingLevel(nr) = 1
'    End Select
'End Sub
'
' 'Inverted Reels
'Sub FadeIReel(nr, reel)
'    Select Case FadingLevel(nr)
'        Case 2:FadingLevel(nr) = 0
'        Case 3:FadingLevel(nr) = 2
'        Case 4:reel.Visible = 1:FadingLevel(nr) = 3
'        Case 5:reel.Visible = 0:FadingLevel(nr) = 1
'    End Select
'End Sub
'
''**********************************************************************************************************
''Digital Display
''**********************************************************************************************************
'
'Dim Digits(32)
'' 1st Player
'Digits(0) = Array(LED10,LED11,LED12,LED13,LED14,LED15,LED16)
'Digits(1) = Array(LED20,LED21,LED22,LED23,LED24,LED25,LED26)
'Digits(2) = Array(LED30,LED31,LED32,LED33,LED34,LED35,LED36)
'Digits(3) = Array(LED40,LED41,LED42,LED43,LED44,LED45,LED46)
'Digits(4) = Array(LED50,LED51,LED52,LED53,LED54,LED55,LED56)
'Digits(5) = Array(LED60,LED61,LED62,LED63,LED64,LED65,LED66)
'Digits(6) = Array(LED70,LED71,LED72,LED73,LED74,LED75,LED76)
'
'' 2nd Player
'Digits(7) = Array(LED80,LED81,LED82,LED83,LED84,LED85,LED86)
'Digits(8) = Array(LED90,LED91,LED92,LED93,LED94,LED95,LED96)
'Digits(9) = Array(LED100,LED101,LED102,LED103,LED104,LED105,LED106)
'Digits(10) = Array(LED110,LED111,LED112,LED113,LED114,LED115,LED116)
'Digits(11) = Array(LED120,LED121,LED122,LED123,LED124,LED125,LED126)
'Digits(12) = Array(LED130,LED131,LED132,LED133,LED134,LED135,LED136)
'Digits(13) = Array(LED140,LED141,LED142,LED143,LED144,LED145,LED146)
'
'' 3rd Player
'Digits(14) = Array(LED150,LED151,LED152,LED153,LED154,LED155,LED156)
'Digits(15) = Array(LED160,LED161,LED162,LED163,LED164,LED165,LED166)
'Digits(16) = Array(LED170,LED171,LED172,LED173,LED174,LED175,LED176)
'Digits(17) = Array(LED180,LED181,LED182,LED183,LED184,LED185,LED186)
'Digits(18) = Array(LED190,LED191,LED192,LED193,LED194,LED195,LED196)
'Digits(19) = Array(LED200,LED201,LED202,LED203,LED204,LED205,LED206)
'Digits(20) = Array(LED210,LED211,LED212,LED213,LED214,LED215,LED216)
'
'' 4th Player
'Digits(21) = Array(LED220,LED221,LED222,LED223,LED224,LED225,LED226)
'Digits(22) = Array(LED230,LED231,LED232,LED233,LED234,LED235,LED236)
'Digits(23) = Array(LED240,LED241,LED242,LED243,LED244,LED245,LED246)
'Digits(24) = Array(LED250,LED251,LED252,LED253,LED254,LED255,LED256)
'Digits(25) = Array(LED260,LED261,LED262,LED263,LED264,LED265,LED266)
'Digits(26) = Array(LED270,LED271,LED272,LED273,LED274,LED275,LED276)
'Digits(27) = Array(LED280,LED281,LED282,LED283,LED284,LED285,LED286)
'
'' Credits
'Digits(28) = Array(LED4,LED2,LED6,LED7,LED5,LED1,LED3)
'Digits(29) = Array(LED18,LED9,LED27,LED28,LED19,LED8,LED17)
'' Balls
'Digits(30) = Array(LED39,LED37,LED48,LED49,LED47,LED29,LED38)
'Digits(31) = Array(LED67,LED58,LED69,LED77,LED68,LED57,LED59)
'
'Sub DisplayTimer_Timer
'	Dim ChgLED,ii,num,chg,stat,obj
'	ChgLed = Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
'If Not IsEmpty(ChgLED) Then
'		If DesktopMode = True Then
'		For ii = 0 To UBound(chgLED)
'			num = chgLED(ii, 0) : chg = chgLED(ii, 1) : stat = chgLED(ii, 2)
'			if (num < 32) then
'				For Each obj In Digits(num)
'					If chg And 1 Then obj.State = stat And 1 
'					chg = chg\2 : stat = stat\2
'				Next
'			else
'			end if
'		next
'		end if
'end if
'End Sub
'
''**********Sling Shot Animations
'' Rstep and Lstep  are the variables that increment the animation
''****************
'Dim RStep, Lstep, RKick, LA, RC, LC, RA, RB, RD
'
'Sub RightSlingShot_Slingshot
'	vpmtimer.pulsesw 36
'    PlaySound SoundFX("right_slingshot",DOFContactors), 0, 1, 0.05, 0.05
'    RSling.Visible = 0
'    RSling1.Visible = 1
'    sling1.TransZ = -20
'    RStep = 0
'    RightSlingShot.TimerEnabled = 1
'End Sub
'
'Sub RightSlingShot_Timer
'    Select Case RStep
'        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.TransZ = -10
'        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.TransZ = 0:RightSlingShot.TimerEnabled = 0
'    End Select
'    RStep = RStep + 1
'End Sub
'
'Sub LeftSlingShot_Slingshot
'	vpmtimer.pulsesw 37
'    PlaySound SoundFX("left_slingshot",DOFContactors),0,1,-0.05,0.05
'    LSling.Visible = 0
'    LSling1.Visible = 1
'    sling2.TransZ = -32
'    LStep = 0
'    LeftSlingShot.TimerEnabled = 1
'End Sub
'
'Sub LeftSlingShot_Timer
'    Select Case LStep
'        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.TransZ = -10
'        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.TransZ = 0:LeftSlingShot.TimerEnabled = 0
'    End Select
'    LStep = LStep + 1
'End Sub
'
'Sub WallLA_Hit
'    RubberLA.Visible = 0
'    RubberLA1.Visible = 1
'    LA = 0
'    WallLA.TimerEnabled = 1
'End Sub
'
'Sub WallLA_Timer
'    Select Case LA
'        Case 3:RubberLA1.Visible = 0:RubberLA2.Visible = 1
'        Case 4:RubberLA2.Visible = 0:RubberLA.Visible = 1:WallLA.TimerEnabled = 0:
'    End Select
'    LA = LA + 1
'End Sub
''
'Sub WallLC_Hit
'    RubberLC.Visible = 0
'    RubberLC1.Visible = 1
'    LC = 0
'    WallLC.TimerEnabled = 1
'End Sub
'
'Sub WallLC_Timer
'    Select Case LC
'        Case 3:RubberLC1.Visible = 0:RubberLC2.Visible = 1
'        Case 4:RubberLC2.Visible = 0:RubberLC.Visible = 1:WallLC.TimerEnabled = 0:
'    End Select
'    LC = LC + 1
'End Sub
''
'Sub WallRA_Hit
'    RubberRA.Visible = 0
'    RubberRA1.Visible = 1
'    RA = 0
'    WallRA.TimerEnabled = 1
'End Sub
'
'Sub WallRA_Timer
'    Select Case RA
'        Case 3:RubberRA1.Visible = 0:RubberRA2.Visible = 1
'        Case 4:RubberRA2.Visible = 0:RubberRA.Visible = 1:WallRA.TimerEnabled = 0:
'    End Select
'    RA = RA + 1
'End Sub
''
'Sub WallRC_Hit
'    RubberRC.Visible = 0
'    RubberRC1.Visible = 1
'    RC = 0
'    WallRC.TimerEnabled = 1
'End Sub
'
'Sub WallRC_Timer
'    Select Case RC
'        Case 3:RubberRC1.Visible = 0:RubberRC2.Visible = 1
'        Case 4:RubberRC2.Visible = 0:RubberRC.Visible = 1:WallRC.TimerEnabled = 0:
'    End Select
'    RC = RC + 1
'End Sub
''
'Sub sw3a_Hit
'	vpmTimer.PulseSw 3
'    RubberDL.Visible = 0
'    RubberDL1.Visible = 1
'    RB = 0
'    sw3a.TimerEnabled = 1
'End Sub
'
'Sub sw3a_Timer
'    Select Case RB
'        Case 3:RubberDL1.Visible = 0:RubberDL2.Visible = 1
'        Case 4:RubberDL2.Visible = 0:RubberDL.Visible = 1:sw3a.TimerEnabled = 0:
'    End Select
'    RB = RB + 1
'End Sub
''
'Sub sw3b_Slingshot
'	vpmTimer.PulseSw 3
'    RubberDR.Visible = 0
'    RubberDR1.Visible = 1
'    RD = 0
'    sw3b.TimerEnabled = 1
'End Sub
'
'Sub sw3b_Timer
'    Select Case RD
'        Case 3:RubberDR1.Visible = 0:RubberDR2.Visible = 1
'        Case 4:RubberDR2.Visible = 0:RubberDR.Visible = 1:sw3b.TimerEnabled = 0:
'    End Select
'    RD = RD + 1
'End Sub
'
'
'Sub RKickTimer_Timer
'    Select Case RKick
'        Case 1:kickarmtop_prim.ObjRotX = -50
'        Case 2:kickarmtop_prim.ObjRotX = -50
'        Case 3:kickarmtop_prim.ObjRotX = -50
'        Case 4:kickarmtop_prim.ObjRotX = -50
'        Case 5:kickarmtop_prim.ObjRotX = -50
'        Case 6:kickarmtop_prim.ObjRotX = -50
'        Case 7:kickarmtop_prim.ObjRotX = -50
'        Case 8:kickarmtop_prim.ObjRotX = -50
'        Case 9:kickarmtop_prim.ObjRotX = -50
'        Case 10:kickarmtop_prim.ObjRotX = -50
'        Case 11:kickarmtop_prim.ObjRotX = -24
'        Case 12:kickarmtop_prim.ObjRotX = -12
'        Case 13:kickarmtop_prim.ObjRotX = 0:RKickTimer.Enabled = 0
'    End Select
'    RKick = RKick + 1
'End Sub
'
'' *********************************************************************
''                      Supporting Ball & Sound Functions
'' *********************************************************************
'
'Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
'    Vol = Csng(BallVel(ball) ^2 / 1000)
'End Function
'
'Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "table1" is the name of the table
'    Dim tmp
'    tmp = ball.x * 2 / table1.width-1
'    If tmp > 0 Then
'        Pan = Csng(tmp ^10)
'    Else
'        Pan = Csng(-((- tmp) ^10) )
'    End If
'End Function
'
'Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
'    Pitch = BallVel(ball) * 20
'End Function
'
'Function BallVel(ball) 'Calculates the ball speed
'    BallVel = INT(SQR((ball.VelX ^2) + (ball.VelY ^2) ) )
'End Function
'
'Sub Pins_Hit (idx)
'	PlaySound "pinhit_low", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0
'End Sub
'
'Sub Targets_Hit (idx)
'	PlaySound "target", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0
'End Sub
'
'Sub Metals_Thin_Hit (idx)
'	PlaySound "metalhit_thin", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
'End Sub
'
'Sub Metals_Medium_Hit (idx)
'	PlaySound "metalhit_medium", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
'End Sub
'
'Sub Metals2_Hit (idx)
'	PlaySound "metalhit2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
'End Sub
'
'Sub Gates_Hit (idx)
'	PlaySound "gate4", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
'End Sub
'
'Sub Spinner_Spin
'	PlaySound "fx_spinner",0,.25,0,0.25
'End Sub
'
'Sub Rubbers_Hit(idx)
' 	dim finalspeed
'  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
' 	If finalspeed > 20 then 
'		PlaySound "fx_rubber2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
'	End if
'	If finalspeed >= 6 AND finalspeed <= 20 then
' 		RandomSoundRubber()
' 	End If
'End Sub
'
'Sub Posts_Hit(idx)
' 	dim finalspeed
'  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
' 	If finalspeed > 16 then 
'		PlaySound "fx_rubber2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
'	End if
'	If finalspeed >= 6 AND finalspeed <= 16 then
' 		RandomSoundRubber()
' 	End If
'End Sub
'
'Sub RandomSoundRubber()
'	Select Case Int(Rnd*3)+1
'		Case 1 : PlaySound "rubber_hit_1", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
'		Case 2 : PlaySound "rubber_hit_2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
'		Case 3 : PlaySound "rubber_hit_3", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
'	End Select
'End Sub
'
'Sub LeftFlipper_Collide(parm)
' 	RandomSoundFlipper()
'End Sub
'
'Sub RightFlipper_Collide(parm)
' 	RandomSoundFlipper()
'End Sub
'
'Sub RandomSoundFlipper()
'	Select Case Int(Rnd*3)+1
'		Case 1 : PlaySound "flip_hit_1", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
'		Case 2 : PlaySound "flip_hit_2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
'		Case 3 : PlaySound "flip_hit_3", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
'	End Select
'End Sub
'
''*****************************************
''      JP's VP10 Rolling Sounds
''*****************************************
'
'Const tnob = 5 ' total number of balls
'ReDim rolling(tnob)
'InitRolling
'
'Sub InitRolling
'    Dim i
'    For i = 0 to tnob
'        rolling(i) = False
'    Next
'End Sub
'
'Sub RollingTimer_Timer()
'    Dim BOT, b
'    BOT = GetBalls
'
'	' stop the sound of deleted balls
'    For b = UBound(BOT) + 1 to tnob
'        rolling(b) = False
'        StopSound("fx_ballrolling" & b)
'    Next
'
'	' exit the sub if no balls on the table
'    If UBound(BOT) = -1 Then Exit Sub
'
'	' play the rolling sound for each ball
'    For b = 0 to UBound(BOT)
'        If BallVel(BOT(b) ) > 1 AND BOT(b).z < 30 Then
'            rolling(b) = True
'            PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) ), Pan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0
'        Else
'            If rolling(b) = True Then
'                StopSound("fx_ballrolling" & b)
'                rolling(b) = False
'            End If
'        End If
'    Next
'End Sub
'
''**********************
'' Ball Collision Sound
''**********************
'
'Sub OnBallBallCollision(ball1, ball2, velocity)
'	PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 2000, Pan(ball1), 0, Pitch(ball1), 0, 0
'End Sub
'
'
'
''************************************
'' What you need to add to your table
''************************************
'
'' a timer called RollingTimer. With a fast interval, like 10
'' one collision sound, in this script is called fx_collide
'' as many sound files as max number of balls, with names ending with 0, 1, 2, 3, etc
'' for ex. as used in this script: fx_ballrolling0, fx_ballrolling1, fx_ballrolling2, fx_ballrolling3, etc
'
'
''******************************************
'' Explanation of the rolling sound routine
''******************************************
'
'' sounds are played based on the ball speed and position
'
'' the routine checks first for deleted balls and stops the rolling sound.
'
'' The For loop goes through all the balls on the table and checks for the ball speed and 
'' if the ball is on the table (height lower than 30) then then it plays the sound
'' otherwise the sound is stopped, like when the ball has stopped or is on a ramp or flying.
'
'' The sound is played using the VOL, PAN and PITCH functions, so the volume and pitch of the sound
'' will change according to the ball speed, and the PAN function will change the stereo position according
'' to the position of the ball on the table.
'
'
''**************************************
'' Explanation of the collision routine
''**************************************
'
'' The collision is built in VP.
'' You only need to add a Sub OnBallBallCollision(ball1, ball2, velocity) and when two balls collide they 
'' will call this routine. What you add in the sub is up to you. As an example is a simple Playsound with volume and paning
'' depending of the speed of the collision.
'
''Bally Power Play
''added by Inkochnito
'Sub editDips
'	Dim vpmDips : Set vpmDips = New cvpmDips
'	With vpmDips
'		.AddForm 700,400,"Power Play - DIP switches"
'		.AddFrame 2,20,190,"Maximum credits",&H00070000,Array("10 credits",&H00010000,"20 credits",&H00030000,"30 credits",&H00050000,"40 credits",&H00070000)'dip 17&18&19
'		.AddFrame 2,99,190,"High game to date award",&H00000060,Array("no award",0,"1 credit",&H00000020,"2 credits",&H00000040,"3 credits",&H00000060)'dip 6&7
'		.AddFrame 2,173,190,"Drop target bank award",&H00002000,Array("2X, 3X, 5X, extra ball, special",0,"2X, 3X, 5X && extra ball, special",&H00002000)'dip 14
'		.AddFrame 2,220,190,"Knocking down 4 targets of 1 bank",&H00200000,Array("will reset both banks",0,"will reset only that bank",&H00200000)'dip 22
'		.AddFrame 205,20,190,"Thumper-bumper adjust",&H00400000,Array("alternating",0,"all bumpers on",&H00400000)'dip 23
'		.AddFrame 205,66,190,"High score feature",&HC0000000,Array("no award",0,"extra ball",&H80000000,"replay",&HC0000000)'dip 31&32
'		.AddFrame 205,127,190,"Left && right alley rollover buttons",&H00004000,Array("scores 100 points",0,"scores 1000 points",&H00004000)'dip 15
'		.AddFrame 205,173,190,"Balls per game", 32768,Array("3 balls",0,"5 balls", 32768)'dip 16
'		.AddFrame 205,220,190,"Outlane special adjust",&H30000000,Array("special does NOT lite",0,"special alternates",&H20000000,"both specials on",&H30000000)'dip 29&30
'		.AddChk 2,0,100,Array("Match feature",&H00100000)'dip 21
'		.AddChk 150,0,100,Array("Credits display",&H00080000)'dip 20
'		.AddChk 295,0,100,Array("Melody option",&H00000080)'dip 8
'		.AddLabel 50,290,300,20,"After hitting OK, press F3 to reset game with new settings."
'		.ViewDips
'	End With
'End Sub
'Set vpmShowDips = GetRef("editDips")
'
'Sub table1_Paused:Controller.Pause = 1:End Sub
'Sub table1_unPaused:Controller.Pause = 0:End Sub
' 
'    Sub table1_Exit:Controller.Stop:End Sub
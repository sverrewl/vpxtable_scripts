'Viking (Bally 1980) v1.0 by bord
'DOF by Arngrim

Option Explicit
Randomize

' Thalamus 2018-07-24
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Moved solenoids above table1_init
' Changed UseSolenoids=1 to 2
' Thalamus 2018-08-28 : Improved directional sounds

' Export targets banks and import them as fx2_droptarget2 and fx2_DTReset2
' Set pos approx 0.67 / -0.24 for the above and -0.82 / -0.43


Const VolDiv = 2000    ' Lower number, louder ballrolling/collition sound
Const VolCol   = 10    ' Ball collition divider ( voldiv/volcol )

' The rest of the values are multipliers
'
'  .5 = lower volume
' 1.5 = higher volume

Const VolBump   = 2    ' Bumpers volume.
Const VolRol    = 1    ' Rollovers volume.
Const VolGates  = 1    ' Gates volume.
Const VolMetal  = 1    ' Metals volume.
Const VolRB     = 1    ' Rubber bands volume.
Const VolRH     = 1    ' Rubber hits volume.
Const VolPo     = 1    ' Rubber posts volume.
Const VolPi     = 1    ' Rubber pins volume.
Const VolPlast  = 1    ' Plastics volume.
Const VolTarg   = 1    ' Targets volume.
Const VolWood   = 1    ' Woods volume.
Const VolKick   = 1    ' Kicker volume.
Const VolSpin   = 1.5  ' Spinners volume.
Const VolFlip   = 1    ' Flipper volume.

Const cGameName = "vikingb"

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "Can't open controller.vbs"
On Error Goto 0

LoadVPM "01130100", "Bally.VBS", 3.21  'Viking
Dim DesktopMode: DesktopMode = table1.ShowDT


'************************************************
'************************************************
'************************************************
'************************************************
'************************************************
Const UseSolenoids = 2
Const UseLamps = 1
Const UseSync = 0

' Standard Sounds
Const SSolenoidOn = "fx_solenoid"
Const SSolenoidOff = "fx_solenoidoff"
Const SCoin = "fx_coin"

Dim bsTrough, dtleft, dtright

Dim nfozzy: nfozzy=1            'set to 0 for standard vpx flipper physics or 1 for nfozzys adjustments to rate of return

'*************nFozzy flipper routine part 1
dim returnspeed, lfstep, rfstep
returnspeed = leftflipper.return
lfstep = 1
rfstep = 1

sub leftflipper_timer()
    select case lfstep
        Case 1: leftflipper.return = returnspeed * 0.6 :lfstep = lfstep + 1
        Case 2: leftflipper.return = returnspeed * 0.7 :lfstep = lfstep + 1
        Case 3: leftflipper.return = returnspeed * 0.8 :lfstep = lfstep + 1
        Case 4: leftflipper.return = returnspeed * 0.9 :lfstep = lfstep + 1
        Case 5: leftflipper.return = returnspeed * 1 :lfstep = lfstep + 1
        Case 6: leftflipper.timerenabled = 0 : lfstep = 1
    end select
end sub

sub rightflipper_timer()
    select case rfstep
        Case 1: rightflipper.return = returnspeed * 0.6 :rfstep = rfstep + 1
        Case 2: rightflipper.return = returnspeed * 0.7 :rfstep = rfstep + 1
        Case 3: rightflipper.return = returnspeed * 0.8 :rfstep = rfstep + 1
        Case 4: rightflipper.return = returnspeed * 0.9 :rfstep = rfstep + 1
        Case 5: rightflipper.return = returnspeed * 1 :rfstep = rfstep + 1
        Case 6: rightflipper.timerenabled = 0 : rfstep = 1
    end select
end sub

'*************nFozzy flipper routine part 1 end

Sub SolLFlipper(Enabled)
    If Enabled Then
        PlaySoundAtVol SoundFX("fx_flipperup",DOFFlippers), LeftFlipper, VolFlip
        LeftFlipper.RotateToEnd
    Else
        PlaySoundAtVol SoundFX("fx_flipperdown",DOFFlippers), LeftFlipper, VolFlip
        LeftFlipper.RotateToStart
'*************nFozzy flipper routine part 2
        if nfozzy=1 then
            LeftFlipper.TimerEnabled = 1
            LeftFlipper.TimerInterval = 16
            LeftFlipper.return = returnspeed * 0.5
        end if
'*************nFozzy flipper routine part 2 end
    End If
End Sub

Sub SolRFlipper(Enabled)
    If Enabled Then
        PlaySoundAtVol SoundFX("fx_flipperup",DOFFlippers), RightFlipper, VolFlip
        RightFlipper.RotateToEnd
   Else
        PlaySoundAtVol SoundFX("fx_flipperdown",DOFFlippers), RightFlipper, VolFlip
        RightFlipper.RotateToStart
'*************nFozzy flipper routine part 2b
        if nfozzy=1 then
            RightFlipper.TimerEnabled = 1
            RightFlipper.TimerInterval = 16
            RightFlipper.return = returnspeed * 0.5
        end if
'*************nFozzy flipper routine part 2b end
    End If
End Sub

Sub FlipperTimer_Timer
    'Add flipper, gate and spinner rotations here
    flipperleft_prim.ObjROTZ = LeftFlipper.CurrentAngle
    flipperright_prim.ObjRoTZ = RightFlipper.CurrentAngle
	batleftshadow.rotz = LeftFlipper.CurrentAngle
	batrightshadow.rotz  = RightFlipper.CurrentAngle
	rampgate_prim.RotY = Gate3.CurrentAngle +34
	rampgate_prim1.RotY = Gate4.CurrentAngle +32
	if sw19.isdropped=1 then dropbank19.visible=0: end if
	if sw19.isdropped=0 then dropbank19.visible=1: end if
	if sw18.isdropped=1 then dropbank18.visible=0: end if
	if sw18.isdropped=0 then dropbank18.visible=1: end if
	if sw17.isdropped=1 then dropbank17.visible=0: end if
	if sw17.isdropped=0 then dropbank17.visible=1: end if
	if sw4.isdropped=0 then dropbank1.visible=0:dropbank2.visible=0:dropbank3.visible=0:dropbank4.visible=1: end if
	if sw3.isdropped=0 then dropbank1.visible=0:dropbank2.visible=0:dropbank3.visible=1:dropbank4.visible=0: end if
	if sw2.isdropped=0 then dropbank1.visible=0:dropbank2.visible=1:dropbank3.visible=0:dropbank4.visible=0: end if
	if sw1.isdropped=0 then dropbank1.visible=1:dropbank2.visible=0:dropbank3.visible=0:dropbank4.visible=0: end if
	if sw1.isdropped=1 then dropbank1.visible=0:dropbank2.visible=0:dropbank3.visible=0:dropbank4.visible=0: end if
	if l59.state=1 then creditlight.image="crediton"
	if l59.state=0 then creditlight.image="creditoff"
End Sub

dim Angle

'***********Rotate Spinner
Sub SpinnerTimer_Timer
	Angle = (sin (sw14.CurrentAngle))
	TextBox1.text = sw14.currentangle
	TextBox2.text = Angle
    SpinnerRod.TransZ = -sin( (sw14.CurrentAngle) * (2*3.14/360)) * 5
    SpinnerRod.TransX = (sin( (sw14.CurrentAngle- 90) * (2*3.14/360)) * -5)
End Sub

'************************************************
' Solenoids
'************************************************
SolCallback(7) =    "bsTrough.SolOut"       'outhole kicker
'SolCallback(2) =   "Solknocker "               'knocker
SolCallback(8) =    "SolKickUpSaucer"       'kick up saucer
SolCallback(9) =    "SolKickDownSaucer"     'kick down saucer
'SolCallback(5) =   ""                  'left slingshot
'SolCallback(6) =   ""                  'right slingshot
SolCallback(1) =    "SolLeftTargetReset "           'in line drop target reset
SolCallback(2) =    "SolRightTargetReset "          '3 drop target reset
'SolCallback(9) =   "vpmSolSound "              'top left thumper bumper
'SolCallback(10) =  "vpmsolsound "              'top right thumper bumper
'SolCallback(11) =  "vpmSolSound "              'left side thumper bumper
'SolCallback(12) =  "vpmSolSound "              'right side thumper bumper
SolCallback(3) =    ""                  '3 drop target 1 (top)
SolCallback(4) =    ""                  '3 drop target 2
SolCallback(5) =    ""                  '3 drop target 3 (bottom)
'SolCallback(16) =  ""                  'coin lockout door
'SolCallback(17) =  ""                  'ki relay (flipper enable)
SolCallback(17) =   "SolTopSaucer"      'top saucer

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub Table1_Init
	vpmInit Me
     With Controller
         .GameName = cGameName
         If Err Then MsgBox "Can't start Game " & cGameName & vbNewLine & Err.Description:Exit Sub
         .SplashInfoLine = "Viking"
         .HandleKeyboard = 0
         .ShowTitle = 0
         .ShowDMDOnly = 1
         .ShowFrame = 0
         .HandleMechanics = False
		.Hidden = 1
'        .Games(cGameName).Settings.Value("rol") = 0 '1= rotated display, 0= normal
        '.SetDisplayPosition 0,0, GetPlayerHWnd 'restore dmd window position
        On Error Resume Next
        Controller.SolMask(0) = 0
        vpmTimer.AddTimer 2000, "Controller.SolMask(0)=&Hffffffff'" 'ignore all solenoids - then add the Timer to renable all the solenoids after 2 seconds
        Controller.Run GetPlayerHWnd
        On Error Goto 0
    End With

    PinMAMETimer.Interval = PinMAMEInterval
    PinMAMETimer.Enabled = 1


'Trough
    Set bsTrough=New cvpmBallStack
    with bsTrough
        .InitSw 0,8,0,0,0,0,0,0
        .InitKick BallRelease,90,7
        .InitExitSnd Soundfx("fx_ballrel",DOFContactors), Soundfx("fx_solenoid",DOFContactors)
        .Balls=1
    end with

     ' Nudging
     vpmNudge.TiltSwitch = 7
     vpmNudge.Sensitivity = 3
     vpmNudge.TiltObj = Array(leftslingshot, rightslingshot, sw37, sw38, sw39, sw40)

    Set dtLeft=New cvpmDropTarget
        dtLeft.InitDrop Array(sw1,sw2,sw3,sw4),Array(1,2,3,4)
        dtLeft.InitSnd SoundFX("fx2_droptarget",DOFDropTargets),SoundFX("fx2_DTReset",DOFDropTargets)

    Set dtRight=New cvpmDropTarget
        dtRight.InitDrop Array(sw17,sw18,sw19),Array(17,18,19)
        dtRight.InitSnd SoundFX("fx2_droptarget2",DOFDropTargets),SoundFX("fx2_DTReset2",DOFDropTargets)

'*****Drop Lights Off

    GILights 1

 End Sub

Sub GILights (enabled)
    Dim light
    For each light in GI:light.State = Enabled: Next
End Sub

'**********************************************************************************************************
'Plunger code
'**********************************************************************************************************

Sub Table1_KeyDown(ByVal KeyCode)
	If KeyDownHandler(keycode) Then Exit Sub
	If keycode = PlungerKey Then Plunger.Pullback:playsoundat"plungerpull", plunger
	if KeyCode = LeftTiltKey Then Nudge 90, 4
	if KeyCode = RightTiltKey Then Nudge 270, 4
	if KeyCode = CenterTiltKey Then Nudge 0, 4
    ' Manual Ball Control
	If keycode = 46 Then	 				' C Key
		If EnableBallControl = 1 Then
			EnableBallControl = 0
		Else
			EnableBallControl = 1
		End If
	End If
    If EnableBallControl = 1 Then
		If keycode = 48 Then 				' B Key
			If BCboost = 1 Then
				BCboost = BCboostmulti
			Else
				BCboost = 1
			End If
		End If
		If keycode = 203 Then BCleft = 1	' Left Arrow
		If keycode = 200 Then BCup = 1		' Up Arrow
		If keycode = 208 Then BCdown = 1	' Down Arrow
		If keycode = 205 Then BCright = 1	' Right Arrow
	End If
End Sub

Sub Table1_KeyUp(ByVal KeyCode)
	If KeyUpHandler(keycode) Then Exit Sub
	If keycode = PlungerKey Then Plunger.Fire:PlaySoundat"plunger", plunger
    'Manual Ball Control
	If EnableBallControl = 1 Then
		If keycode = 203 Then BCleft = 0	' Left Arrow
		If keycode = 200 Then BCup = 0		' Up Arrow
		If keycode = 208 Then BCdown = 0	' Down Arrow
		If keycode = 205 Then BCright = 0	' Right Arrow
	End If
End Sub

Sub ShooterLane_Hit
End Sub

Sub BallRelease_UnHit
End Sub

Sub Drain_Hit()
    PlaySoundAt "fx2_drain2", Drain : bstrough.addball me
End Sub

Sub sw24_Hit   : Controller.Switch(24) = True : PlaySoundAt "fx_hole-enter", sw24:End Sub
'
Sub sw32_Hit   : Controller.Switch(32) = True : PlaySoundAt "fx_hole-enter", sw32:End Sub

dim k1step, k2step, k3step

Sub SolTopSaucer(enabled)
    if enabled then
        PlaySoundAtVol Soundfx("fx_ballrel",DOFContactors), sw24, VolKick
        PlaySoundAtVol Soundfx("fx_solenoid",DOFContactors), sw24, VolKick
        Controller.Switch(24) = false
        sw24.Kick  177, 18 + 5 * Rnd
		cupun2.rotx = -25
		k3step=0
		kicktoptimer.enabled=1
    end if
End Sub

Sub kicktoptimer_Timer
    Select Case K3Step
        Case 3:cupun2.rotx = -25
        Case 4:cupun2.rotx = -25
        Case 5:cupun2.rotx = -25
        Case 6:cupun2.rotx = -25
        Case 7:cupun2.rotx = -25
        Case 8:cupun2.rotx = -18
        Case 9:cupun2.rotx = -11
        Case 10:cupun2.rotx = 5
        Case 11:cupun2.rotx = 0:KickToptimer.Enabled = 0
    End Select
    k3Step = k3Step + 1
End Sub

Sub SolKickUpSaucer(enabled)
    if enabled then
        PlaySoundAtVol Soundfx("fx_ballrel",DOFContactors), sw32, VolKick
        PlaySoundAtVol Soundfx("fx_solenoid",DOFContactors), sw32, VolKick
        Controller.Switch(32) = false
        sw32.Kick  10, 24 + 5 * Rnd
		cupdeux2.rotx = 25
		k1step = 0
		KickUptimer.Enabled = 1
    end if
End Sub

Sub kickuptimer_Timer
    Select Case K1Step
        Case 3:cupdeux2.rotx = 25
        Case 4:cupdeux2.rotx = 25
        Case 5:cupdeux2.rotx = 25
        Case 6:cupdeux2.rotx = 25
        Case 7:cupdeux2.rotx = 25
        Case 8:cupdeux2.rotx = 18
        Case 9:cupdeux2.rotx = 11
        Case 10:cupdeux2.rotx = 5
        Case 11:cupdeux2.rotx = 0:KickUptimer.Enabled = 0
    End Select
    k1Step = k1Step + 1
End Sub

Sub SolKickDownSaucer(enabled)
    if enabled then
        PlaySoundAtVol Soundfx("fx_ballrel",DOFContactors), sw32, VolKick
        PlaySoundAtVol Soundfx("fx_solenoid",DOFContactors), sw32, VolKick
        Controller.Switch(32) = false
        sw32.Kick  180, 18 + 5 * Rnd
		cupdeux2.rotx = -25
		k2step = 0
		KickDowntimer.Enabled=1
    end if
End Sub

Sub kickdowntimer_Timer
    Select Case K2Step
        Case 3:cupdeux2.rotx = -25
        Case 4:cupdeux2.rotx = -25
        Case 5:cupdeux2.rotx = -25
        Case 6:cupdeux2.rotx = -25
        Case 7:cupdeux2.rotx = -25
        Case 8:cupdeux2.rotx = -18
        Case 9:cupdeux2.rotx = -11
        Case 10:cupdeux2.rotx = -5
        Case 11:cupdeux2.rotx = 0:KickDowntimer.Enabled = 0
    End Select
    k2Step = k2Step + 1
End Sub

'Drop Targets
 Sub Sw1_Dropped:dtLeft.Hit 1 : End Sub
 Sub Sw2_Dropped:dtLeft.Hit 2 : End Sub
 Sub Sw3_Dropped:dtLeft.Hit 3 : End Sub
 Sub Sw4_Dropped:dtLeft.Hit 4 : End Sub

 Sub Sw17_Dropped:dtRight.Hit 1 : End Sub
 Sub Sw18_Dropped:dtRight.Hit 2 : End Sub
 Sub Sw19_Dropped:dtRight.Hit 3 : End Sub

Sub SolRightTargetReset(enabled)
    dim xx
    if enabled then
        dtRight.SolDropUp enabled
    end if
End Sub

Sub SolLeftTargetReset(enabled)
    dim xx
    if enabled then
        dtLeft.SolDropUp enabled
    end if
End Sub

'Sub DropTimer_Timer
'	If GIisoff=true then dropplate1.visible=0
'	If GIisoff=true then dropplate2.visible=0
'	If GIisoff=true then dropplate3.visible=0
'	If GIisoff=false and sw49.isdropped=1 then dropplate1.visible=0
'	If GIisoff=false and sw49.isdropped=0 then dropplate1.visible=1
'	If GIisoff=false and sw50.isdropped=1 then dropplate2.visible=0
'	If GIisoff=false and sw50.isdropped=0 then dropplate2.visible=1
'	If GIisoff=false and sw51.isdropped=1 then dropplate3.visible=0
'	If GIisoff=false and sw51.isdropped=0 then dropplate3.visible=1
'End Sub

'Bumpers

Sub sw37_Hit : vpmTimer.PulseSw 37 : PlaySoundAtVol SoundFX("fx2_bumper_1",DOFContactors), sw37, VolBump: End Sub
Sub sw38_Hit : vpmTimer.PulseSw 38 : PlaySoundAtVol SoundFX("fx2_bumper_1",DOFContactors), sw38, VolBump: End Sub
Sub sw39_Hit : vpmTimer.PulseSw 39 : PlaySoundAtVol SoundFX("fx2_bumper_2",DOFContactors), sw39, VolBump: End Sub
Sub sw40_Hit : vpmTimer.PulseSw 40 : PlaySoundAtVol SoundFX("fx2_bumper_3",DOFContactors), sw40, VolBump: End Sub

'Wire Triggers
Sub SW12_Hit:Controller.Switch(12)=1 : End Sub  'B
Sub SW12_unHit:Controller.Switch(12)=0:End Sub
Sub SW13_Hit:Controller.Switch(13)=1 : End Sub  'A
Sub SW13_unHit:Controller.Switch(13)=0:End Sub
Sub SW15_Hit:Controller.Switch(15)=1 : End Sub  'Side R.O. Button
Sub SW15_unHit:Controller.Switch(15)=0:End Sub
Sub SW26_Hit:Controller.Switch(26)=1 : End Sub  'Right out Rollover
Sub SW26_unHit:Controller.Switch(26)=0:End Sub
Sub SW27_Hit:Controller.Switch(27)=1 : End Sub  'Flip Feed Lane (Rt)
Sub SW27_unHit:Controller.Switch(27)=0:End Sub
Sub SW28_Hit:Controller.Switch(28)=1 : End Sub  'Flip Feed Lane (Lt)
Sub SW28_unHit:Controller.Switch(28)=0:End Sub
Sub SW29_Hit:Controller.Switch(29)=1 : End Sub  'Left outlane
Sub SW29_unHit:Controller.Switch(29)=0:End Sub
Sub SW30_Hit:Controller.Switch(30)=1 : End Sub  'Right Side Lane R.O.
Sub SW30_unHit:Controller.Switch(30)=0:End Sub

'Spinners
Sub sw14_Spin : vpmTimer.PulseSw (14) :PlaySoundAtVol "fx_spinner", sw14, VolSpin: End Sub

'Targets
Sub sw5_Hit:vpmTimer.PulseSw (5):End Sub
Sub sw25_Hit:vpmTimer.PulseSw (25):End Sub

Sub SolKnocker(Enabled)
    If Enabled Then PlaySound SoundFX("Fx_Knocker",DOFKnocker)
End Sub

'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep

Sub RightSlingShot_Slingshot
    PlaySoundAt SoundFX("fx_slingshot",DOFContactors), sling1
    vpmtimer.PulseSw(35)
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.ObjRotY = 25
    RStep = 0
    RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.ObjRotY = 11
        Case 4:RSLing2.Visible = 0:RSLing3.Visible = 1:sling1.ObjRotY = 0
        Case 5:RSLing3.Visible = 0:RSLing.Visible = 1:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
    PlaySoundAt SoundFX("fx_slingshot",DOFContactors), sling2
    vpmtimer.pulsesw(36)
    LSling.Visible = 0
    LSling1.Visible = 1
    sling2.ObjRotY = 27
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.ObjRotY = 11
        Case 4:LSLing2.Visible = 0:LSLing3.Visible = 1:sling2.ObjRotY = 0
        Case 5:LSLing3.Visible = 0:LSLing.Visible = 1:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

'-------------------------------------
' Map lights into array
' Set unmapped lamps to Nothing
'-------------------------------------
Set Lights(1)  = l1
Set Lights(2)  = l2
Set Lights(3)  = l3
Set Lights(4)  = l4
Set Lights(5)  = l5
Set Lights(6)  = l6
Set Lights(7)  = l7
Set Lights(8)  = l8
Set Lights(9)  = l9
Lights(10) = Array(l10,l10a)
'Set Lights(11) = l11 'Shoot Again
Set Lights(12) = l12
'Set Lights(13) = l13 'Ball In Play
Set Lights(14) = l14
Set Lights(15) = l15
'Set Lights(16) = unused
Set Lights(17) = l17
Set Lights(18) = l18
Set Lights(19) = l19
Set Lights(20) = l20
Set Lights(21) = l21
Set Lights(22) = l22
Set Lights(23) = l23
Set Lights(24) = l24
Set Lights(25) = l25
Set Lights(26) = l26
'Set Lights(27) = l27 'LightMatch
Set Lights(28) = l28
'Set Lights(29) = l29 'LightHighScore
Set Lights(30) = l30
Set Lights(31) = l31
'Set Lights(32) = unused
Set Lights(33) = l33
Set Lights(34) = l34
Set Lights(35) = l35
Set Lights(36) = l36
Set Lights(37) = l37
Set Lights(38) = l38
Set Lights(39) = l39
Set Lights(40) = l40
Set Lights(41) = l41
Set Lights(42) = l42
Set Lights(43) = l43
Set Lights(44) = l44
'Set Lights(45) = l45 'LightGameOver
Set Lights(46) = l46
Set Lights(47) = l47
'Set Lights(48) = unused
Set Lights(49) = l49
Set Lights(50) = l50
Set Lights(51) = l51
Set Lights(52) = l52
Set Lights(53) = l53
Set Lights(54) = l54
Set Lights(55) = l55
Set Lights(56) = l56
Set Lights(57) = l57
Set Lights(58) = l58
Set Lights(59) = l59 'LightCredit
Set Lights(60) = l60
'Set Lights(61) = l61 'LightTilt
Set Lights(62) = l62
Set Lights(63) = l63

'Set Lights(10) = GILeft        'PF_Left
'Set Lights(42) = GIRight   'PF_Right
'Set Lights(26) = GICenter  'PF_Center

'---------------------------------------------------------------
' Edit the dip switches
'---------------------------------------------------------------
Sub editDips
    Dim vpmDips : Set vpmDips = New cvpmDips
    with vpmDips
        .AddForm 315, 370, "Viking DIP Switch Settings"
        .AddFrame 0,0,190,"Maximum credits",&H03000000,Array("10 credits",0,"20 credits",&H01000000,"30 credits",&H02000000,"40 credits",&H03000000)'dip 25&26
        .AddFrame 0,76,190,"High game to date",&H00300000,Array("no award",0,"1 credit",&H00100000,"2 credits",&H00200000,"3 credits",&H00300000)'dip 21&22
        .AddFrame 0,152,190,"Bumper points adjust",&H00800000,Array("100 points",0,"1.000 points",&H00800000)'dip 24
        .AddFrame 0,244,190,"Red target adjust",&H00000080,Array("does not add bonus",0,"adds 5 extra bonus",&H00000080)'dip 8
        .AddFrame 205,0,190,"Sound features",&H30000000,Array("chime effects",0,"noises and no background",&H10000000,"noise effects",&H20000000,"noises and background",&H30000000)'dip 29&30
        .AddFrame 205,76,190,"High score feature",&H00000060,Array("points",0,"extra ball",&H00000040,"replay",&H00000060)'dip 6&7
        .AddChk 205,155,190,Array("Match feature",&H08000000)'dip 28
		.AddChk 205,175,190,Array("Credits displayed",&H04000000)'dip 27
        .AddChk 205,195,190,Array("25K light in memory",&H00004000)'dip 15
        .AddChk 205,215,190,Array("3 bank drop target in memory",32768)'dip 16
        .AddChk 205,235,190,Array("Special and lock ball in memory",&H00002000)'dip 14
        .AddChk 205,255,190,Array("In-line extra ball && special in memory",&H80000000)'dip 32
        .AddChk 205,275,190,Array("In-line drop targets points in memory",&H00400000)'dip 23
        .AddLabel 50,300,300,20,"After hitting OK, press F3 to reset game with new settings."
        .ViewDips
    end with
End Sub
Set vpmShowDips = GetRef("editDips")

Sub table1_Paused:Controller.Pause = 1:End Sub
Sub table1_unPaused:Controller.Pause = 0:End Sub
Sub table1_Exit:Controller.Stop:End Sub

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
  PlaySound sound, 1, Vol, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

'Set all as per ball position & speed, but Vol Multiplier may be used eg; PlaySoundAtBallVol "sound",3

Sub PlaySoundAtBallVol(sound, VolMult)
  PlaySound sound, 0, Vol(ActiveBall) * VolMult, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 1, AudioFade(ActiveBall)
End Sub

'Set position as bumperX and Vol manually.

Sub PlaySoundAtBumperVol(sound, tableobj, Vol)
  PlaySound sound, 1, Vol, AudioPan(tableobj), 0,0,1, 1, AudioFade(tableobj)
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
		StopSound("fx_Rolling_Wood" & b)
		StopSound("fx_Rolling_Plastic" & b)
		StopSound("fx_Rolling_Metal" & b)
	Next

	' exit the sub if no balls on the table
	If UBound(BOT) = -1 Then Exit Sub

	' play the rolling sound for each ball
	For b = 0 to UBound(BOT)
		If BallVel(BOT(b) ) > 1 Then

			' ***Ball on WOOD playfield***
			if BOT(b).z < 27 Then
				PlaySound("fx_Rolling_Wood" & b), -1, Vol(BOT(b) )/2, AudioPan(BOT(b) )/5, 0, Pitch(BOT(b) ), 1, 0, AudioFade(BOT(b) )
				rolling(b) = True
				StopSound("fx_Rolling_Plastic" & b)
				StopSound("fx_Rolling_Metal" & b)
			Else
				if rolling(b) = true Then
					rolling(b) = False
					StopSound("fx_Rolling_Wood" & b)
					StopSound("fx_Rolling_Plastic" & b)
					StopSound("fx_Rolling_Metal" & b)
				end if
			End If
		Else
			If rolling(b) = True Then
				StopSound("fx_Rolling_Wood" & b)
				StopSound("fx_Rolling_Plastic" & b)
				StopSound("fx_Rolling_Metal" & b)
				rolling(b) = False
			End If
		End If

		'***Ball Drop Sounds***

		If BOT(b).VelZ < -1 and BOT(b).z < 50 and BOT(b).z > 27 Then 'height adjust for ball drop sounds
			PlaySoundAtBOTBallZ "fx_ball_drop" & b, BOT(b)
		End If

		debug.print BOT(b).x & " " & BOT(b).y

		if SQR((BOT(b).VelX ^2) + (BOT(b).VelY ^2)) < 0.2 and InRect(BOT(b).x, BOT(b).y, 420,290,440,290, 440, 300, 420,300) Then
			BOT(b).vely = 5
		end if
	Next
End Sub

Sub PlaySoundAtBOTBall(sound, BOT)
		PlaySound sound, 0, Vol(BOT), AudioPan(BOT), 0, Pitch(BOT), 0, 1, AudioFade(BOT)
End Sub

Sub PlaySoundAtBOTBallZ(sound, BOT)
		PlaySound sound, 0, ABS(BOT.velz)/17, AudioPan(BOT), 0, Pitch(BOT), 1, 0, AudioFade(BOT)
End Sub

'*** Determines if a Points (px,py) is inside a 4 point polygon A-D in Clockwise/CCW order
Function InRect(px,py,ax,ay,bx,by,cx,cy,dx,dy)
	Dim AB, BC, CD, DA
	AB = (bx*py) - (by*px) - (ax*py) + (ay*px) + (ax*by) - (ay*bx)
	BC = (cx*py) - (cy*px) - (bx*py) + (by*px) + (bx*cy) - (by*cx)
	CD = (dx*py) - (dy*px) - (cx*py) + (cy*px) + (cx*dy) - (cy*dx)
	DA = (ax*py) - (ay*px) - (dx*py) + (dy*px) + (dx*ay) - (dy*ax)

	If (AB <= 0 AND BC <=0 AND CD <= 0 AND DA <= 0) Or (AB >= 0 AND BC >=0 AND CD >= 0 AND DA >= 0) Then
		InRect = True
	Else
		InRect = False
	End If
End Function


'**********************
' Ball Collision Sound
'**********************

Sub OnBallBallCollision(ball1, ball2, velocity)
	PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 2000, Pan(ball1), 0, Pitch(ball1), 0, 0
End Sub

'*****************************************
'	ninuzzu's	BALL SHADOW
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
            BallShadow(b).X = ((BOT(b).X) - (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/7)) + 6
        Else
            BallShadow(b).X = ((BOT(b).X) + (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/7)) - 6
        End If
        ballShadow(b).Y = BOT(b).Y + 12
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
	PlaySound "pinhit_low", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Targets_Hit (idx)
	PlaySound "target", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Metals_Thin_Hit (idx)
	PlaySound "metalhit_thin", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Metals_Medium_Hit (idx)
	PlaySound "metalhit_medium", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Metals2_Hit (idx)
	PlaySound "metalhit2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Gates_Hit (idx)
	PlaySound "gate4", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Spinner_Spin
	PlaySound "fx_spinner", 0, .25, AudioPan(Spinner), 0.25, 0, 0, 1, AudioFade(Spinner)
End Sub

Sub Rubbers_Hit(idx)
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 20 then
		PlaySound "fx_rubber2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End if
	If finalspeed >= 6 AND finalspeed <= 20 then
 		RandomSoundRubber()
 	End If
End Sub

Sub Posts_Hit(idx)
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 16 then
		PlaySound "fx_rubber2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End if
	If finalspeed >= 6 AND finalspeed <= 16 then
 		RandomSoundRubber()
 	End If
End Sub

Sub RandomSoundRubber()
	Select Case Int(Rnd*3)+1
		Case 1 : PlaySound "rubber_hit_1", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 2 : PlaySound "rubber_hit_2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 3 : PlaySound "rubber_hit_3", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
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
		Case 1 : PlaySound "flip_hit_1", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 2 : PlaySound "flip_hit_2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 3 : PlaySound "flip_hit_3", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End Select
End Sub

Dim EnableBallControl
EnableBallControl = false 'Change to true to enable manual ball control (or press C in-game) via the arrow keys and B (boost movement) keys

'*****************************************
'   rothbauerw's Manual Ball Control
'*****************************************

Dim BCup, BCdown, BCleft, BCright
Dim ControlBallInPlay, ControlActiveBall
Dim BCvel, BCyveloffset, BCboostmulti, BCboost

BCboost = 1				'Do Not Change - default setting
BCvel = 4				'Controls the speed of the ball movement
BCyveloffset = -0.01 	'Offsets the force of gravity to keep the ball from drifting vertically on the table, should be negative
BCboostmulti = 3		'Boost multiplier to ball veloctiy (toggled with the B key)

ControlBallInPlay = false

Sub StartBallControl_Hit()
	Set ControlActiveBall = ActiveBall
	ControlBallInPlay = true
End Sub

Sub StopBallControl_Hit()
	ControlBallInPlay = false
End Sub

Sub BallControlTimer_Timer()
	If EnableBallControl and ControlBallInPlay then
		If BCright = 1 Then
			ControlActiveBall.velx =  BCvel*BCboost
		ElseIf BCleft = 1 Then
			ControlActiveBall.velx = -BCvel*BCboost
		Else
			ControlActiveBall.velx = 0
		End If

		If BCup = 1 Then
			ControlActiveBall.vely = -BCvel*BCboost
		ElseIf BCdown = 1 Then
			ControlActiveBall.vely =  BCvel*BCboost
		Else
			ControlActiveBall.vely = bcyveloffset
		End If
	End If
End Sub

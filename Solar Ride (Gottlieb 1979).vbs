Option Explicit
Randomize

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Const cGameName="solaride",UseSolenoids=1,UseLamps=1,UseGI=0,SSolenoidOn="SolOn",SSolenoidOff="SolOff",SFlipperOn="fx_Flipperup",SFlipperOff="fx_Flipperdown"
Const SCoin="coin",cCredits=""

LoadVPM"01150000","GTS1.VBS",3.22

'**********************************************************
'********   	OPTIONS		*******************************
'**********************************************************

Dim BallShadows: Ballshadows=0  		'******************set to 1 to turn on Ball shadows
Dim FlipperShadows: FlipperShadows=1  '***********set to 1 to turn on Flipper shadows
Dim FlipperColor: FlipperColor=0		'************* set to 0 for yellow, 1 for red, or 2 for white flipper Rubbers


'*************************************************************

'Solenoid Call backs
'**********************************************************************************************************
SolCallback(1)= "DrainKick"	
SolCallback(2)=  "vpmSolSound SoundFX(""Knocker"",DOFKnocker),"
SolCallback(6)= "SolSaucerRight"
SolCallBack(7)= "SolSaucerUpper"
SolCallback(8)= "Lraised"
SolCallback(17)= "FastFlips.TiltSol"
'SolCallback(17)="vpmNudge.SolGameOn"
SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub Lraised(enabled)
	if enabled then Lreset.enabled=True
End sub

Sub Lreset_timer
	for each objekt in droptargets
		objekt.isdropped=0
	next
	controller.switch(30) = 0
	controller.switch(31) = 0
	controller.switch(32) = 0
	controller.switch(33) = 0
	controller.switch(34) = 0
	PlaySoundAt SoundFX("DTreset", DOFContactors), sw30
	for each Light in DTlights:light.state=0: next
	lreset.enabled=False
end sub

Sub DrainKick(enabled)
	If enabled Then
	Drain.kick 65, 10
	PlaySoundAt SoundFX("ballrelease",DOFContactors), Drain
	End If
end Sub

Sub Drain_Hit
	PlaySoundAt "drain", Drain
	controller.switch(66) = 1
End Sub

sub Drain_unhit
	controller.Switch(66) = 0
end sub

Sub SolLFlipper(Enabled)
     If Enabled Then
		PlaySound SoundFX("fx_flipperup",DOFFlippers), 1, 1.6, AudioPan(LFLIPsound), 0,0,0, 1, AudioFade(LFLIPsound)
		LeftFlipper.RotateToEnd
		LeftFlipper1.RotateToEnd
     Else
		PlaySound SoundFX("fx_flipperdown",DOFFlippers), 1, 1.6, AudioPan(LFLIPsound), 0,0,0, 1, AudioFade(LFLIPsound)
		LeftFlipper.RotateToStart
		LeftFlipper1.RotateToStart
     End If
  End Sub
  
Sub SolRFlipper(Enabled)
     If Enabled Then
		PlaySoundAt SoundFX("fx_flipperup",DOFFlippers), RightFlipper
		RightFlipper.RotateToEnd
     Else
		PlaySoundAt SoundFX("fx_flipperdown",DOFFlippers), RightFlipper
		RightFlipper.RotateToStart
     End If
End Sub

Sub SolSaucerRight(Enabled)
    If Enabled Then
        sw41.kick 265,6
        Controller.Switch(41) = False
        PlaySoundAt SoundFX("popper_ball", DOFContactors), sw41
        sw41.uservalue=1
        sw41.timerenabled=1
        PkickarmR.rotz=15
    end if
End Sub
 
Sub sw41_timer
    select case sw41.uservalue
      case 2:
        PkickarmR.rotz=0
        me.timerenabled=0
    end Select
    sw41.uservalue=sw41.uservalue+1
End Sub
 
Sub SolSaucerUpper(Enabled)
    If Enabled Then
        Sw42.kick 180,7
        Controller.Switch(42) = False
        PlaySoundAt SoundFX("popper_ball", DOFContactors), sw42
        sw42.uservalue=1
        sw42.timerenabled=1
        Pkickarm.rotz=15
    end if
End Sub
 
Sub sw42_timer
    select case sw42.uservalue
      case 2:
        Pkickarm.rotz=0
        me.timerenabled=0
    end Select
    sw42.uservalue=sw42.uservalue+1
End Sub

'**********************************************************************************************************

'Initiate Table
'**********************************************************************************************************
Dim bsTrough, dtDropL, FastFlips
Dim objekt, BPG, hsaward

Sub SolarRide_Init
	vpmInit Me
	On Error Resume Next
		With Controller
		.GameName = cGameName
		If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
		.SplashInfoLine = "Solar Ride (Gotlieb 1979)"
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

	Set FastFlips = new cFastFlips
	with FastFlips
		.CallBackL = "SolLflipper"	'Point these to flipper subs
		.CallBackR = "SolRflipper"	'...
'		.CallBackUL = "SolULflipper"'...(upper flippers, if needed)
'		.CallBackUR = "SolURflipper"'...
		.TiltObjects = True 'Optional, if True calls vpmnudge.solgameon automatically. IF YOU GET A LINE 1 ERROR, DISABLE THIS! (or setup vpmNudge.TiltObj!)
		.InitDelay "FastFlips", 100			'Optional, if > 0 adds some compensation for solenoid jitter (occasional problem on Bram Stoker's Dracula)
		.DebugOn = False		'Debug, always-on flippers. Call FastFlips.DebugOn True or False in debugger to enable/disable.
	end with

	vpmNudge.TiltSwitch = 4
	vpmNudge.Sensitivity = 3
	vpmNudge.TiltObj = Array(Bumper1,Bumper2, Bumper3,RightSlingshot)


	FindDips

	If B2SOn Then 'Show Desktop components
		for each objekt in BackdropStuff: objekt.visible=0: next
	  Else
		for each objekt in BackdropStuff: objekt.visible=1: next
	End if

	if ballshadows=1 then
		BallShadowUpdate.enabled=1
	  else
		BallShadowUpdate.enabled=0
	end if

	if flippershadows=1 then 
		FlipperLSh.visible=1
		FlipperLSh1.visible=1
		FlipperRSh.visible=1
	  else
		FlipperLSh.visible=0
		FlipperLSh1.visible=0
		FlipperRSh.visible=0
	end if

	if FlipperColor=0 then
		FlipperT1.image="flipper_gottlieb_leftYellow"
		FlipperT2.image="flipper_gottlieb_leftYellow"
		FlipperT5.image="flipper_gottlieb_rightYellow"
	end if

	if FlipperColor=1 then
		FlipperT1.image="flipper_gottlieb_leftRed"
		FlipperT2.image="flipper_gottlieb_leftRed"
		FlipperT5.image="flipper_gottlieb_rightRed"
	end if

	if FlipperColor=2 then
		FlipperT1.image="flipper_gottlieb_leftWhite"
		FlipperT2.image="flipper_gottlieb_leftWhite"
		FlipperT5.image="flipper_gottlieb_rightWhite"
	end if

	drain.createball
    startGame.enabled=true

End Sub

Sub SolarRide_Paused:Controller.Pause = 1:End Sub
 
Sub SolarRide_unPaused:Controller.Pause = 0:End Sub
 
Sub SolarRide_Exit
    If b2son then controller.stop
End Sub

sub startGame_timer
    dim xx
    PlaySoundAt "poweron", Plunger
    LampTimer.enabled=1
    For each xx in GI:xx.State = 1: Next        '*****GI Lights On
    me.enabled=false
end sub

'**********************************************************************************************************
'Plunger code
'**********************************************************************************************************

Sub SolarRide_KeyDown(ByVal KeyCode)
	If keycode = LeftFlipperKey Then FastFlips.FlipL True :  FastFlips.FlipUL True
	If keycode = RightFlipperKey Then FastFlips.FlipR True :  FastFlips.FlipUR True
	If vpmKeyDown(KeyCode) Then Exit Sub
	If keycode = PlungerKey Then Plunger.Pullback:PlaySoundAt"plungerpull", Plunger
	If keycode=AddCreditKey then PlaySoundAt "coin", Drain: vpmTimer.pulseSW (swCoin1): end if
End Sub

Sub SolarRide_KeyUp(ByVal KeyCode)
	If keycode = LeftFlipperKey Then FastFlips.FlipL False :  FastFlips.FlipUL False
   	If keycode = RightFlipperKey Then FastFlips.FlipR False :  FastFlips.FlipUR False

	If keycode = 61 then FindDips
	If vpmKeyUp(KeyCode) Then Exit Sub
	If keycode = PlungerKey Then Plunger.Fire:PlaySoundAt"plunger", Plunger
End Sub

'**********************************************************************************************************

'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep

Sub RightSlingShot_Slingshot
	vpmTimer.PulseSw 13
	DOF 203, DOFPulse
    PlaySoundAt SoundFXDOF("right_slingshot",203,DOFPulse,DOFContactors), slingR
    RSling.Visible = 0
    RSling1.Visible = 1
	slingR.objroty = -15	
    RStep = 0
    RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:slingR.objroty = -7
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:slingR.objroty = 0:RightSlingShot.TimerEnabled = 0:
    End Select
    RStep = RStep + 1
End Sub

'Switches
'**********************************************************************************************************
'**********************************************************************************************************
'Drain kickers

Sub sw41_Hit()
	controller.Switch(41) = True
	PlaySoundAt "popper_ball", sw41
End Sub  

Sub sw42_Hit()
	controller.Switch(42) = True
	PlaySoundAt "popper_ball", sw42
End Sub  

'Bumpers
Sub Bumper1_Hit 
	vpmTimer.PulseSw(40)
	PlaySoundAt SoundFX("fx_bumper1",DOFContactors), Bumper1
End Sub

Sub Bumper2_Hit
	vpmTimer.PulseSw(43) 
	PlaySoundAt SoundFXDOF("fx_bumper1",202, DOFPulse, DOFContactors), Bumper2
End Sub

Sub Bumper3_Hit 
	vpmTimer.PulseSw(43)
	PlaySoundAt SoundFXDOF("fx_bumper1",201, DOFPulse, DOFContactors), Bumper3
End Sub

'Drop Targets
Sub sw30_Dropped:controller.switch(30) = 1:Lsw30.state=1:End Sub 
Sub sw31_Dropped:controller.switch(31) = 1:Lsw31.state=1:End Sub 
Sub sw32_Dropped:controller.switch(32) = 1:Lsw32.state=1:End Sub 
Sub sw33_Dropped:controller.switch(33) = 1:Lsw33.state=1:End Sub 
Sub sw34_Dropped:controller.switch(34) = 1:Lsw34.state=1:End Sub

'Star Triggers
Sub sw12a_Hit():Controller.Switch(12)=1 :End Sub
Sub sw12a_UnHit:Controller.Switch(12)=0:End Sub
Sub sw12b_Hit():Controller.Switch(12)=1 : End Sub  
Sub sw12b_UnHit:Controller.Switch(12)=0:End Sub 
Sub sw12c_Hit():Controller.Switch(12)=1 : End Sub 
Sub sw12c_UnHit:Controller.Switch(12)=0:End Sub

'Scoring Gate

 Sub Gate2_Hit:vpmTimer.PulseSw(52): End Sub

'Wire Triggers
Sub sw10a_Hit():Controller.Switch(10)=1 : DOF 204, DOFOn:End Sub  
Sub sw10a_UnHit:Controller.Switch(10)=0:DOF 204, DOFOff:End Sub 
Sub sw10b_Hit():Controller.Switch(10)=1 : DOF 205, DOFOn:End Sub  
Sub sw10b_UnHit:Controller.Switch(10)=0:DOF 205, DOFOff:End Sub
Sub sw11_Hit():Controller.Switch(11)=1 :  End Sub  
Sub sw11_UnHit:Controller.Switch(11)=0:End Sub
Sub sw20_Hit():Controller.Switch(20)=1 : End Sub  
Sub sw20_UnHit:Controller.Switch(20)=0:End Sub 
Sub sw21_Hit():Controller.Switch(21)=1 : End Sub 
Sub sw21_UnHit:Controller.Switch(21)=0:End Sub 
Sub sw22_Hit():Controller.Switch(22)=1 : End Sub
Sub sw22_UnHit:Controller.Switch(22)=0:End Sub 
Sub sw23_Hit():Controller.Switch(23)=1 : End Sub 
Sub sw23_UnHit:Controller.Switch(23)=0:End Sub 

'**********************************************************
'*********** Animated scoring Rubbers
'**********************************************************
 


Sub sw13a_Hit()
	vpmTimer.PulseSw(13)
	RubberA.visible=0
	RubberA1.visible=1
	me.uservalue=1
	me.timerenabled=1
End Sub

sub sw13a_timer
	select case me.uservalue
		case 1: rubbera1.visible=0: rubberA.visible=1
		case 2: rubbera.visible=0: rubbera2.visible=1
		case 3: rubbera2.visible=0: rubbera.visible=1: me.timerenabled=0
	end Select
	me.uservalue=me.uservalue+1
end sub

Sub sw13b_Hit()
	vpmTimer.PulseSw(13)
	rubberb.visible=0
	rubberb1.visible=1
	me.uservalue=1
	me.timerenabled=1
End Sub

sub sw13b_timer
	select case me.uservalue
		case 1: rubberb1.visible=0: rubberb.visible=1
		case 2: rubberb.visible=0: rubberb2.visible=1
		case 3: rubberb2.visible=0: rubberb.visible=1: me.timerenabled=0
	end Select
	me.uservalue=me.uservalue+1
end sub

Sub sw13c_Hit()
	vpmTimer.PulseSw(13)
	rubberc.visible=0
	rubberc1.visible=1
	me.uservalue=1
	me.timerenabled=1
End Sub

sub sw13c_timer
	select case me.uservalue
		case 1: rubberc1.visible=0: rubberc.visible=1
		case 2: rubberc.visible=0: rubberc2.visible=1
		case 3: rubberc2.visible=0: rubberc.visible=1: me.timerenabled=0
	end Select
	me.uservalue=me.uservalue+1
end sub

Sub sw13d_Hit()
	vpmTimer.PulseSw(13)
	rubberd.visible=0
	rubberd1.visible=1
	me.uservalue=1
	me.timerenabled=1
End Sub

sub sw13d_timer
	select case me.uservalue
		case 1: rubberd1.visible=0: rubberd.visible=1
		case 2: rubberd.visible=0: rubberd2.visible=1
		case 3: rubberd2.visible=0: rubberd.visible=1: me.timerenabled=0
	end Select
	me.uservalue=me.uservalue+1
end sub

Sub sw13e_Hit()
	vpmTimer.PulseSw(13)
	rubbere.visible=0
	rubbere1.visible=1
	me.uservalue=1
	me.timerenabled=1
End Sub

sub sw13e_timer
	select case me.uservalue
		case 1: rubbere1.visible=0: rubbere.visible=1
		case 2: rubbere.visible=0: rubbere2.visible=1
		case 3: rubbere2.visible=0: rubbere.visible=1: me.timerenabled=0
	end Select
	me.uservalue=me.uservalue+1
end sub

Sub sw13f_Hit():vpmTimer.PulseSw(13) : End Sub


'**********************************************************
'*********** Lights
'**********************************************************

Dim N1, Light
N1=0
Set LampCallback = GetRef("UpdateMultipleLamps")
 
Sub UpdateMultipleLamps


	N1=Controller.Lamp(1) 'Game Over triggers match and BIP
		If N1 then 
			EMReelBIP.setvalue 1
			EMReelNTM.setvalue 0
			EMReelGO.setvalue 0
		  else
			EMReelBIP.setvalue 0
			EMReelNTM.setvalue 1
			EMReelGO.setvalue 1
		end if

	N1=Controller.Lamp(2) 'Tilt
		If N1 then
			EMReelTilt.setvalue 1
		  else
			EMReelTilt.setvalue 0
		end if

	N1=Controller.Lamp(4) 'Shoot Again
		If N1 then
			EMReelSPSA.setvalue 1
		  else
			EMReelSPSA.setvalue 0
		end if

	N1=Controller.Lamp(3) 'HIGH SCORE TO DATE
		if N1 then
			EMReelHTD.setvalue 1
		  else
			EMReelHTD.setvalue 0
		end if

 End Sub
 


'**********************************************************************************************************
 
'Map lights to an array
'**********************************************************************************************************
'Set Lights(1) = l1 'Game Over Backbox
'Set Lights(2) = l2 'Tilt Backbox
'Set Lights(3) = l3 'High Score Backbox
Set Lights(4) = l4 ' Shoot Again Playfied and backglass
Set Lights(5) = l5
Set Lights(6) = l6
Set Lights(7) = l7
Set Lights(8) = l8
Set Lights(9) = l9
Set Lights(10) = l10
Set Lights(11) = l11
Set Lights(12) = l12
Set Lights(13) = l13
Set Lights(14) = l14
Set Lights(15) = l15
Set Lights(16) = l16
'Set Lights(17) = l17
Set Lights(18) = l18
'Set Lights(19) = l19
'Set Lights(20) = l20
'Set Lights(21) = l21
'Set Lights(22) = l22
Set Lights(23) = l23
Set Lights(24) = l24
Set Lights(25) = l25
Set Lights(26) = l26
Set Lights(27) = l27
Set Lights(28) = l28
Set Lights(29) = l29
Set Lights(30) = l30
Set Lights(31) = l31
Set Lights(32) = l32
Set Lights(33) = l33
'Set Lights(34) = l34
'Set Lights(35) = L35
'Set Lights(36) = l36



'***************************************************************
'*******************	DESKTOP BACKDROP

Dim Digits(32)
 
'Score displays
 
Digits(0)=Array(a1,a2,a3,a4,a5,a6,a7,n,a8)
Digits(1)=Array(a9,a10,a11,a12,a13,a14,a15,n,a16)
Digits(2)=Array(a17,a18,a19,a20,a21,a22,a23,n,a24)
Digits(3)=Array(a25,a26,a27,a28,a29,a30,a31,n,a32)
Digits(4)=Array(a33,a34,a35,a36,a37,a38,a39,n,a40)
Digits(5)=Array(a41,a42,a43,a44,a45,a46,a47,n,a48)
Digits(6)=Array(a49,a50,a51,a52,a53,a54,a55,n,a56)
Digits(7)=Array(a57,a58,a59,a60,a61,a62,a63,n,a64)
Digits(8)=Array(a65,a66,a67,a68,a69,a70,a71,n,a72)
Digits(9)=Array(a73,a74,a75,a76,a77,a78,a79,n,a80)
Digits(10)=Array(a81,a82,a83,a84,a85,a86,a87,n,a88)
Digits(11)=Array(a89,a90,a91,a92,a93,a94,a95,n,a96)
Digits(12)=Array(a97,a98,a99,a100,a101,a102,a103,n,a104)
Digits(13)=Array(a105,a106,a107,a108,a109,a110,a111,n,a112)
Digits(14)=Array(a113,a114,a115,a116,a117,a118,a119,n,a120)
Digits(15)=Array(a121,a122,a123,a124,a125,a126,a127,n,a128)
Digits(16)=Array(a129,a130,a131,a132,a133,a134,a135,n,a136)
Digits(17)=Array(a137,a138,a139,a140,a141,a142,a143,n,a144)
Digits(18)=Array(a145,a146,a147,a148,a149,a150,a151,n,a152)
Digits(19)=Array(a153,a154,a155,a156,a157,a158,a159,n,a160)
Digits(20)=Array(a161,a162,a163,a164,a165,a166,a167,n,a168)
Digits(21)=Array(a169,a170,a171,a172,a173,a174,a175,n,a176)
Digits(22)=Array(a177,a178,a179,a180,a181,a182,a183,n,a184)
Digits(23)=Array(a185,a186,a187,a188,a189,a190,a191,n,a192)
 
'Ball in Play and Credit displays
 
Digits(26)=Array(e00,e01,e02,e03,e04,e05,e06,n)
Digits(27)=Array(e10,e11,e12,e13,e14,e15,e16,n)
Digits(24)=Array(f00,f01,f02,f03,f04,f05,f06,n)
Digits(25)=Array(f10,f11,f12,f13,f14,f15,f16,n)
 
 
Sub DisplayTimer_Timer
    Dim ChgLED,ii,num,chg,stat, obj
    ChgLed = Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
    If Not IsEmpty(ChgLED) Then
        If not b2son Then
        For ii = 0 To UBound(chgLED)
            num = chgLED(ii, 0) : chg = chgLED(ii, 1) : stat = chgLED(ii, 2)
            if (num < 28 ) then
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
 
'Finding an individual dip state based on scapino's Strikes and spares dip code - from unclewillys pinball pool, added another section to get high score award to set replay cards
Dim TheDips(32)
  Sub FindDips
    Dim DipsNumber
	DipsNumber = Controller.Dip(1)
	TheDips(16) = Int(DipsNumber/128)
	If TheDips(16) = 1 then DipsNumber = DipsNumber - 128 end if
	TheDips(15) = Int(DipsNumber/64)
	If TheDips(15) = 1 then DipsNumber = DipsNumber - 64 end if
	TheDips(14) = Int(DipsNumber/32)
	If TheDips(14) = 1 then DipsNumber = DipsNumber - 32 end if
	TheDips(13) = Int(DipsNumber/16)
	If TheDips(13) = 1 then DipsNumber = DipsNumber - 16 end if
	TheDips(12) = Int(DipsNumber/8)
	If TheDips(12) = 1 then DipsNumber = DipsNumber - 8 end if
	TheDips(11) = Int(DipsNumber/4)
	If TheDips(11) = 1 then DipsNumber = DipsNumber - 4 end if
	TheDips(10) = Int(DipsNumber/2)
	If TheDips(10) = 1 then DipsNumber = DipsNumber - 2 end if
	TheDips(9) = Int(DipsNumber)
	DipsNumber = Controller.Dip(2)
	TheDips(24) = Int(DipsNumber/128)
	If TheDips(24) = 1 then DipsNumber = DipsNumber - 128 end if
	TheDips(23) = Int(DipsNumber/64)
	If TheDips(23) = 1 then DipsNumber = DipsNumber - 64 end if
	TheDips(22) = Int(DipsNumber/32)
	If TheDips(22) = 1 then DipsNumber = DipsNumber - 32 end if
	TheDips(21) = Int(DipsNumber/16)
	If TheDips(21) = 1 then DipsNumber = DipsNumber - 16 end if
	TheDips(20) = Int(DipsNumber/8)
	If TheDips(20) = 1 then DipsNumber = DipsNumber - 8 end if
	TheDips(19) = Int(DipsNumber/4)
	If TheDips(19) = 1 then DipsNumber = DipsNumber - 4 end if
	TheDips(18) = Int(DipsNumber/2)
	If TheDips(18) = 1 then DipsNumber = DipsNumber - 2 end if
	TheDips(17) = Int(DipsNumber)
	DipsTimer.Enabled=1
 End Sub


 Sub DipsTimer_Timer()
	hsaward = TheDips(22)
	BPG = TheDips(9)
	dim ebplay: ebplay= TheDips(11)
	If BPG = 1 then 
		if ebplay = 1 then
			instcard.image="InstCard3Balls"
		  else
			instcard.image="InstCard3BallsEB"
		end if
	  Else
		if ebplay = 1 then
			instcard.image="InstCard5Balls"
		  else
			instcard.image="InstCard5BallsEB"
		end if
	End if
	if ebplay = 1 then
		repcard.image="replaycard"&hsaward
	  else
		repcard.image="replaycardeb"&hsaward
	end if
	DipsTimer.enabled=0
 End Sub


'**********************************************************************************************************
'**********************************************************************************************************
Sub editDips
   Dim vpmDips : Set vpmDips = New cvpmDips
   With vpmDips
  .AddForm 700,400,"System 1 - DIP switches"
  .AddFrame 0,0,190,"Coin chutecontrol",&H00040000,Array("seperate",0,"same",&H00040000)'dip 19
  .AddFrame 0,46,190,"Game mode",&H00000400,Array("extraball",0,"replay",&H00000400)'dip 11
  .AddFrame 0,92,190,"High game to date awards",&H00200000,Array("noaward",0,"3 replays",&H00200000)'dip 22
  .AddFrame 0,138,190,"Balls per game",&H00000100,Array("5 balls",0,"3balls",&H00000100)'dip 9
  .AddFrame 0,184,190,"Tilt effect",&H00000800,Array("game over",0,"ball in play only",&H00000800)'dip 12
  .AddFrame 205,0,190,"Maximum credits",&H00030000,Array("5 credits",0,"8 credits",&H00020000,"10 credits",&H00010000,"15 credits",&H00030000)'dip 17&18
  .AddFrame 205,76,190,"Soundsettings",&H80000000,Array("sounds",0,"tones",&H80000000)'dip 32
  .AddFrame 205,122,190,"Attract tune",&H10000000,Array("no attract tune",0,"attract tune played every 6 minutes",&H10000000)'dip 29
  .AddChk 205,175,190,Array("Match feature",&H00000200)'dip 10
  .AddChk 205,190,190,Array("Credits displayed",&H00001000)'dip 13
  .AddChk 205,205,190,Array("Play credit button tune",&H00002000)'dip 14
  .AddChk 205,220,190,Array("Play tones when scoring",&H00080000)'dip 20
  .AddChk 205,235,190,Array("Play coin switch tune",&H00400000)'dip 23
  .AddChk 205,250,190,Array("High game to date displayed",&H00100000)'dip 21
  .AddLabel 50,280,300,20,"After hitting OK, press F3 to reset game with new settings."
  .ViewDips
  End With
 End Sub
 Set vpmShowDips = GetRef("editDips")
 


'*********************************************************************
'                 Positional Sound Playback Functions
'*********************************************************************

' Play a sound, depending on the X,Y position of the table element (especially cool for surround speaker setups, otherwise stereo panning only)
' parameters (defaults): loopcount (1), volume (1), randompitch (0), pitch (0), useexisting (0), restart (1))
' Note that this will not work (currently) for walls/slingshots as these do not feature a simple, single X,Y position
Sub PlayXYSound(soundname, tableobj, loopcount, volume, randompitch, pitch, useexisting, restart)
	PlaySound soundname, loopcount, volume, AudioPan(tableobj), randompitch, pitch, useexisting, restart, AudioFade(tableobj)
End Sub

' Similar subroutines that are less complicated to use (e.g. simply use standard parameters for the PlaySound call)
Sub PlaySoundAt(soundname, tableobj)
    PlaySound soundname, 1, 1, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

Sub PlaySoundAtBall(soundname)
    PlaySoundAt soundname, ActiveBall
End Sub


'*********************************************************************
'                     Supporting Ball & Sound Functions
'*********************************************************************

Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "table1" is the name of the table
	Dim tmp
    tmp = tableobj.y * 2 / SolarRide.height-1
    If tmp > 0 Then
		AudioFade = Csng(tmp ^10)
    Else
        AudioFade = Csng(-((- tmp) ^10) )
    End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "table1" is the name of the table
    Dim tmp
    tmp = tableobj.x * 2 / SolarRide.width-1
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
        StopSound("fx_ballrolling" & b)
    Next

	' exit the sub if no balls on the table
    If UBound(BOT) = -1 Then Exit Sub

	' play the rolling sound for each ball
    For b = 0 to UBound(BOT)
        If BallVel(BOT(b) ) > 1 AND BOT(b).z < 30 Then
            rolling(b) = True
            PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b)), AudioPan(BOT(b)), 0, Pitch(BOT(b)), 1, 0, AudioFade(BOT(b))
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
	PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 2000, AudioPan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub


'*****************************************
'			FLIPPER SHADOWS
'*****************************************

sub FlipperTimer_Timer()
	FlipperT1.roty = LeftFlipper.currentangle
	FlipperT2.roty = LeftFlipper1.currentangle
	FlipperT5.roty = RightFlipper.currentangle
	Pgate.rotz = (Gate.currentangle*.75)+25

	if FlipperShadows=1 then
		FlipperLSh.RotZ = LeftFlipper.currentangle
		FlipperLSh1.RotZ = LeftFlipper.currentangle
		FlipperRSh.RotZ = RightFlipper.currentangle
	end if

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
        If BOT(b).X < SolarRide.Width/2 Then
            BallShadow(b).X = ((BOT(b).X) - (Ballsize/16) + ((BOT(b).X - (SolarRide.Width/2))/17))' + 13
        Else
            BallShadow(b).X = ((BOT(b).X) + (Ballsize/16) + ((BOT(b).X - (SolarRide.Width/2))/17))' - 13
        End If
        ballShadow(b).Y = BOT(b).Y + 10
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


Sub Triggers_Hit (idx)
	playsound "rollover", 0,1,AudioPan(ActiveBall),0,0,0,1,AudioFade(ActiveBall)
End sub

Sub Pins_Hit (idx)
	PlaySound "pinhit_low", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Targets_Hit (idx)
	PlaySound "target", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub DropTargets_Hit (idx)
	PlaySound "DTDrop", 0, 1, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
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
	If finalspeed >= 3 AND finalspeed <= 16 then
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

Sub LeftFlipper1_Collide(parm)
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

'cFastFlips by nFozzy
'Bypasses pinmame callback for faster and more responsive flippers
'Version 1.1 beta2 (More proper behaviour, extra safety against script errors)

'Flipper / game-on Solenoid # reference
'Atari: Sol16
'Astro:  ?
'Bally Early 80's: Sol19
'Bally late 80's (Blackwater 100, etc): Sol19
'Game Plan: Sol16
'Gottlieb System 1: Sol17
'Gottlieb System 80: No dedicated flipper solenoid? GI circuit Sol10?
'Gottlieb System 3: Sol32
'Playmatic: Sol8
'Spinball: Sol25
'Stern (80's): Sol19
'Taito: ?
'Williams System 3, 4, 6: Sol23
'Williams System 7: Sol25
'Williams System 9: Sol23
'Williams System 11: Sol23
'Bally / Williams WPC 90', 92', WPC Security: Sol31
'Data East (and Sega pre-whitestar): Sol23
'Zaccaria: ???

'********************Setup*******************:

'....somewhere outside of any subs....
'dim FastFlips

'....table init....
'Set FastFlips = new cFastFlips
'with FastFlips
'	.CallBackL = "SolLflipper"	'Point these to flipper subs
'	.CallBackR = "SolRflipper"	'...
''	.CallBackUL = "SolULflipper"'...(upper flippers, if needed)
''	.CallBackUR = "SolURflipper"'...
'	.TiltObjects = True 'Optional, if True calls vpmnudge.solgameon automatically. IF YOU GET A LINE 1 ERROR, DISABLE THIS! (or setup vpmNudge.TiltObj!)
''	.InitDelay "FastFlips", 100			'Optional, if > 0 adds some compensation for solenoid jitter (occasional problem on Bram Stoker's Dracula)
''	.DebugOn = False		'Debug, always-on flippers. Call FastFlips.DebugOn True or False in debugger to enable/disable.
'end with

'...keydown section... commenting out upper flippers is not necessary as of 1.1
'If KeyCode = LeftFlipperKey then FastFlips.FlipL True :  FastFlips.FlipUL True
'If KeyCode = RightFlipperKey then FastFlips.FlipR True :  FastFlips.FlipUR True
'(Do not use Exit Sub, this script does not handle switch handling at all!)

'...keyUp section...
'If KeyCode = LeftFlipperKey then FastFlips.FlipL False :  FastFlips.FlipUL False
'If KeyCode = RightFlipperKey then FastFlips.FlipR False :  FastFlips.FlipUR False

'...Solenoid...
'SolCallBack(31) = "FastFlips.TiltSol"
'//////for a reference of solenoid numbers, see top /////


'One last note - Because this script is super simple it will call flipper return a lot.
'It might be a good idea to add extra conditional logic to your flipper return sounds so they don't play every time the game on solenoid turns off
'Example:
'Instead of
		'LeftFlipper.RotateToStart
		'playsound SoundFX("FlipperDown",DOFFlippers), 0, 1, 0.01	'return
'Add Extra conditional logic:
		'LeftFlipper.RotateToStart
		'if LeftFlipper.CurrentAngle = LeftFlipper.StartAngle then
		'	playsound SoundFX("FlipperDown",DOFFlippers), 0, 1, 0.01	'return
		'end if
'That's it]
'*************************************************

Function NullFunction(aEnabled):End Function	'1 argument null function placeholder
Class cFastFlips
	Public TiltObjects, DebugOn, hi
	Private SubL, SubUL, SubR, SubUR, FlippersEnabled, Delay, LagCompensation, Name, FlipState(3)
	
	Private Sub Class_Initialize()
		Delay = 0 : FlippersEnabled = False : DebugOn = False : LagCompensation = False
		Set SubL = GetRef("NullFunction"): Set SubR = GetRef("NullFunction") : Set SubUL = GetRef("NullFunction"): Set SubUR = GetRef("NullFunction")
	End Sub
	
	'set callbacks
	Public Property Let CallBackL(aInput)  : Set SubL  = GetRef(aInput) : Decouple sLLFlipper, aInput: End Property
	Public Property Let CallBackUL(aInput) : Set SubUL = GetRef(aInput) : End Property
	Public Property Let CallBackR(aInput)  : Set SubR  = GetRef(aInput) : Decouple sLRFlipper, aInput:  End Property
	Public Property Let CallBackUR(aInput) : Set SubUR = GetRef(aInput) : End Property
	Public Sub InitDelay(aName, aDelay) : Name = aName : delay = aDelay : End Sub	'Create Delay
	'Automatically decouple flipper solcallback script lines (only if both are pointing to the same sub) thanks gtxjoe
	Private Sub Decouple(aSolType, aInput)  : If StrComp(SolCallback(aSolType),aInput,1) = 0 then SolCallback(aSolType) = Empty End If : End Sub

	'call callbacks
	Public Sub FlipL(aEnabled)
		FlipState(0) = aEnabled	'track flipper button states: the game-on sol flips immediately if the button is held down (1.1)
		If not FlippersEnabled and not DebugOn then Exit Sub
		subL aEnabled
	End Sub

	Public Sub FlipR(aEnabled)
		FlipState(1) = aEnabled
		If not FlippersEnabled and not DebugOn then Exit Sub
		subR aEnabled
	End Sub

	Public Sub FlipUL(aEnabled)
		FlipState(2) = aEnabled
		If not FlippersEnabled and not DebugOn then Exit Sub
		subUL aEnabled
	End Sub	

	Public Sub FlipUR(aEnabled)
		FlipState(3) = aEnabled
		If not FlippersEnabled and not DebugOn then Exit Sub
		subUR aEnabled
	End Sub	
	
	Public Sub TiltSol(aEnabled)	'Handle solenoid / Delay (if delayinit)

		If delay > 0 and not aEnabled then 	'handle delay
			vpmtimer.addtimer Delay, Name & ".FireDelay" & "'"
			LagCompensation = True
		else
			If Delay > 0 then LagCompensation = False
			EnableFlippers(aEnabled)
		end If
	End Sub
	
	Sub FireDelay() : If LagCompensation then EnableFlippers False End If : End Sub
	
	Private Sub EnableFlippers(aEnabled)
		If aEnabled then SubL FlipState(0) : SubR FlipState(1) : subUL FlipState(2) : subUR FlipState(3)
		FlippersEnabled = aEnabled
		If TiltObjects then vpmnudge.solgameon aEnabled
		If Not aEnabled then
			subL False
			subR False
			If not IsEmpty(subUL) then subUL False
			If not IsEmpty(subUR) then subUR False
		End If		
	End Sub
	
End Class
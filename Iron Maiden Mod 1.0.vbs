
'############################################################################################
'#######                                                                             ########
'#######                       FirePower   (Drop Target Mod                          ########
'#######                            (Williams 1980)                                  ########
'#######                                                                             ########
'############################################################################################

' Version 1.0 DarthMarino, Based on WED21's table
'
' Thanks To:
' GTXJoe for the Primative Collection
' Walamab for script help from his FirePower
' UncleReamus and Noah Fentz for some images from their VP9 FirePower
' Hauntfreaks for help in the planet redraw!
' Flupper for the physics starting point
' Every other author for their amazing work!

'Iron Maiden Mod by Siggi 2019

Option Explicit
Randomize

' Thalamus 2019 March : Improved directional sounds
' !! NOTE : Table not verified yet !!

' Options
' Volume devided by - lower gets higher sound

Const VolDiv = 2000    ' Lower number, louder ballrolling/collition sound
Const VolCol = 10      ' Ball collition divider ( voldiv/volcol )

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

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Const cGameName="frpwr_b7",UseLamps=1,UseGI=0,SSolenoidOn="SolOn",SSolenoidOff="SolOff", SCoin="coin"

Const UseSolenoids=25 'FastFlips for system6

Const BallSize = 52
Const Ballmass = 1.8

LoadVPM "01560000", "S6.VBS", 3.26

Dim T1Step,T2Step,T3Step,T4Step,T5Step,T6Step,topPWRStep,midPWRStep,BtmPWRStep,CTStep,DTMod
Dim trTrough


'-----------------------------------------------
'***********************************************
'-----------------------------------------------
'
' T O G G L E   D R O P   T A R G E T   M O D   O P T I O N
'-----------------------------------------------------------
'
'To play with the drop target mod, set the below line to DTMod=1
'To play with the production targets, set the below line to DTMod-0
'-----------------------------------------------
'***********************************************
'-----------------------------------------------

DTMod=1

'---------------------------------------------------
' Set up consts with easy names for solenoid numbers
Const cBallRelease = 1
Const cLeftEjectHole = 4
Const cRightEjectHole = 5
Const cUpperEjectHole = 6
Const cBallSaveKick = 7
Const cBallRampThrower = 8
Const cCreditKnocker = 14
Const cFlashLamps = 15
Const cTopLeftBumper = 17
Const cBottomLeftBumper = 18
Const cTopRightBumper = 19
Const cBottomRightBumper = 20
Const cRightSlingShot = 21
Const cLeftSlingShot = 22
'---------------------------------------------------

'---------------------------------------------------
' Setup Consts with easy names for switch numbers
Const cOutHoleSW = 9
Const cLeftOutsideRolloverSW = 10
Const cLeftInsideRolloverSW = 11
Const cLeftKickerSW = 12
Const cLeftEjectHoleSW = 13
Const cUpperMiddleLeftStandupSW = 14
Const cSpinnerSW = 15
Const cTopLeftStandupSW = 16
Const cTarget1SW = 17
Const cTarget2SW = 18
Const cTarget3SW = 19
Const cTarget4SW = 21
Const cTarget5SW = 22
Const cTarget6SW = 23
Const cBottomLeftBumperSW = 25
Const cTopLeftBumperSW = 26
Const cTopRightBumperSW = 27
Const cBottomRightBumperSW = 28
Const cTopCenterTargetSW = 29
Const cRightEjectHoleSW = 30
Const cUpperTopRightStandupSW = 31
Const cFRolloverSW = 32
Const cIRolloverSW = 33
Const cRRolloverSW = 34
Const cERolloverSW = 35
Const cUpperRightEjectHoleSW = 36
Const cLowerTopRightStnadupSW = 37
Const cMIddleRightStandupSW = 38
Const cTopPOWERTargetSW = 39
Const cMiddlePOWERTargetSW = 40
Const cBottomPOWERTargetSW = 41
Const cRightKickerSW = 42
Const cRightInsideRolloverSW = 43
Const cRightOutsideRolloverSW = 44
Const cRightFlipperSW = 45
Const cBallShooterSW = 46
Const cPlayfieldTiltSW = 47
Const cLowerRightStandupSW = 48
Const cCenterMiddleLeftStandupSW = 49
Const cLowerMiddleLeftStuandupSW = 50
Const cLeftBallRampSW = 51
Const cLeftEjectRolloverSW = 53
Const cRightEjectRolloverSW = 54
Const cRightBallRampSW = 57
Const cCenterBallRampSW = 58
'--------------------------------------------------

'Solenoids Setup
SolCallback(cBallRelease) = "trTrough.SolIn"
SolCallback(cLeftEjectHole) = "LeftEjectHole"
SolCallback(cRightEjectHole) = "RightEjectHole"
SolCallback(cUpperEjectHole) = "UpperEjectHole"
SolCallback(cBallSaveKick) = "BallSaveKick"
SolCallback(cBallRampThrower) = "trTrough.SolOut"
SolCallback(cFlashLamps) = "Flashers"
SolCallback(cCreditKnocker) = "CreditKnocker"
SolCallback(cRightSlingShot) = "sRightSlingShot"
SolCallback(cLeftSlingShot) = "sLeftSlingShot"
SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"
If dtmod=1 then
     SolCallback(2) = "DTLbank.soldropup"
     SolCallback(3) = "DTRbank.soldropup"
End If

'------------------------------------------------------------------
' Setup Solenoid Subs
Sub LeftEjectHole(enabled)
	if enabled Then
		KLeftEjectHole.Kick 180, 15
		controller.Switch(cLeftEjectHoleSW)=0
		Playsound "Popper_ball",0, 1*3, 0.1, 1
	end If
End Sub

Sub RightEjectHole(enabled)
	if enabled Then
		KRightEjectHole.Kick 180,15
		Controller.Switch(cRightEjectHoleSW)=0
		Playsound "Popper_ball",0, 1*3, 0.1, 0.0
	end If
End Sub

Sub UpperEjectHole(enabled)
	if enabled Then
		KUpperEjectHole.Kick -90,15
		Controller.Switch(cUpperRightEjectHoleSW)=0
	Playsound "Popper_ball",0, 1*3, 0.1, 0.5
	end if
End Sub

Sub CreditKnocker(enabled)
	if enabled then
	Playsound "knocker",0, 1, 0.1, 0.5
	'DOF 14,2
	end if
end Sub


'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep

Sub RightSlingShot_Slingshot()
	vpmTimer.PulseSW cRightKickerSW
End Sub

Sub LeftSlingShot_Slingshot()
	vpmTimer.PulseSW cLeftKickerSW
End Sub

Sub sRightSlingShot(enabled)
	If enabled Then
    PlaySoundAtVol "slingshotRight", sling1, 1
    Playsound "IM_maiden_scream"
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.TransZ = -20
    RStep = 0
    RightSlingShot.TimerEnabled = 1
	End If
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.TransZ = -10
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.TransZ = 0:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

Sub sLeftSlingShot(enabled)
	If enabled Then
    PlaySoundAtVol "slingshotLeft", sling2, 1
    Playsound "IM_maiden_scream"
    LSling.Visible = 0
    LSling1.Visible = 1
    sling2.TransZ = -20
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
	End If
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.TransZ = -10
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.TransZ = 0:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

'Drop Targets
dim dtLbank, dtRbank

         set dtLBank = new cvpmdroptarget
         With dtLBank
             .InitDrop Array(DTarget1, DTarget2, DTarget3), Array(17, 18, 19)
             .Initsnd "droptargetL", "resetdropL"
             .CreateEvents "dtLBank"
         End With

         set dtRBank = new cvpmdroptarget
         With dtRBank
             .InitDrop Array(DTarget4, DTarget5, DTarget6), Array(21, 22, 23)
             .Initsnd "droptargetR", "resetdropR"
             .CreateEvents "dtRBank"
         End With

Sub LBankReset(enabled)
 		If enabled Then
 			dtLBank.SolDropUp enabled
 		End If
 	End Sub

 	Sub RBankReset(enabled)
 		If enabled Then
 			dtRBank.SolDropUp enabled
 		End If
 End Sub

'Drop Target Reset Safeguards
Sub LTcheck_Timer()
If DTarget1.isdropped=1 and DTarget2.isdropped=1 and DTarget3.isdropped=1  then
LTReset.Enabled=True
LTCheck.enabled=False
End If
End Sub
Sub RTcheck_Timer()
If DTarget4.isdropped=1 and DTarget5.isdropped=1 and DTarget6.isdropped=1 then
RTReset.Enabled=True
RTCheck.enabled=False
End If
End Sub

Sub LTReset_Timer()
If DTarget1.isdropped=1 and DTarget2.isdropped=1 and DTarget3.isdropped=1  then
Dtarget1.isdropped=0
DTarget2.isdropped=0
DTarget3.isdropped=0
playsound "resetdropL"
End If
LTCheck.Enabled=True
LTReset.Enabled=False
End Sub

Sub RTReset_Timer()
If DTarget4.isdropped=1 and DTarget5.isdropped=1 and DTarget6.isdropped=1 then
DTarget4.isdropped=0
DTarget5.isdropped=0
DTarget6.isdropped=0
playsound "resetdropR"
End If
RTCheck.Enabled=True
RTReset.Enabled=False
End Sub


'-----------------------------------------------------------------------
' Flipper Solenoid Handlers

sub SolRFlipper(enabled)
	if enabled Then
		PlaySoundAtVol "Fx_FlipperUp", RightFlipper, 1
		RightFlipper.RotateToEnd

	Else
		PlaySoundAtVol "Fx_FlipperDown", RightFlipper, 1
		RightFlipper.RotateToStart
	end If
end Sub


sub SolLFlipper(enabled)
	if enabled Then
		PlaySoundAtVol "Fx_FlipperUp", LeftFlipper, 1
		LeftFlipper.RotateToEnd
	Else
		PlaySoundAtVol "Fx_FlipperDown", LeftFlipper, 1
		LeftFlipper.RotateToStart
	end If
end Sub
'-----------------------------------------------------------------------
' Handle eject hole hit and unhit events and set corresponding switches
Sub KLeftEjectHole_Hit()
	controller.Switch(cLeftEjectHoleSW)=1
    PlaysoundAtVol "IM_no6", ActiveBall, 1
End Sub

Sub KLeftEjectHole_UnHit()
	controller.Switch(cLeftEjectHoleSW)=0
End Sub

Sub KRightEjectHole_Hit()
	controller.Switch(cRightEjectHoleSW)=1
    PlaysoundAtVol "IM_we_want", ActiveBall, 1
End Sub

Sub KRightEjectHole_UnHit()
	controller.Switch(cRightEjectHoleSW)=0
End Sub

Sub KUpperEjectHole_Hit()
	controller.Switch(cUpperRightEjectHoleSW)=1
    PlaysoundAtVol "IM_free_man", ActiveBall, 1
End Sub

Sub KUpperEjectHole_Unhit()
	controller.Switch(cUpperRightEjectHoleSW)=0
End Sub

Sub BallSaveKick(enabled)
	if enabled Then
		Playsound "slingshotleft", 0, .67, -0.05, 0.05
	End If

End Sub

Sub Flashers(enabled)
	If enabled Then
		LFire1.State = LightStateOn
		LFire2.State = LightStateOn
		LPower1.State = LightStateOn
		LPower2.State = LightStateOn
	Else
		LFire1.State = LightStateOff
		LFire2.State = LightStateOff
		LPower1.State = LightStateOff
		LPower2.State = LightStateOff
	End If
End Sub
'--------------------------------------------------------------------------------------
' MUSIC PLAYER

Dim musicNum
Sub music_hit()
    If musicNum = 0 then PlayMusic "Iron Maiden - 22 Acacia Avenue.mp3" End If
	If musicNum = 1 then PlayMusic "Iron Maiden - Aces High.mp3" End If
    If musicNum = 2 then PlayMusic "Iron Maiden - Children Of The Damned.mp3" End If
    If musicNum = 3 then PlayMusic "Iron Maiden - The Wicker Man.mp3" End If
    If musicNum = 4 then PlayMusic "Iron Maiden - Rainmaker.mp3" End If
    If musicNum = 5 then PlayMusic "Iron Maiden - The Trooper.mp3" End If
    'If musicNum = 6 then PlayMusic "Iron Maiden - Revelations.mp3" End If
    'If musicNum = 7 then PlayMusic "Iron Maiden - The Prisoner.mp3" End If
    'If musicNum = 8 then PlayMusic "Iron Maiden - Wrathchild.mp3" End If
    'If musicNum = 9 then PlayMusic "Iron Maiden - Hallowed Be Thy Name.mp3" End If
    'If musicNum = 10 then PlayMusic "Iron Maiden - Stranger In A Strange Land.mp3" End If
    'If musicNum = 11 then PlayMusic "Iron Maiden - Phantom Of The Opera.mp3" End If
    'If musicNum = 12 then PlayMusic "Iron Maiden - Wasted Years.mp3" End If
    'If musicNum = 13 then PlayMusic "Iron Maiden - Where Eagles Dare.mp3" End If
    'If musicNum = 14 then PlayMusic "Iron Maiden - Charlotte The Harlot.mp3" End If
    'If musicNum = 15 then PlayMusic "Iron Maiden - Flight Of Icarus.mp3" End If
    'If musicNum = 16 then PlayMusic "Iron Maiden - Transylvania.mp3" End If
    'If musicNum = 17 then PlayMusic "Iron Maiden - Two Minutes To Midnight.mp3" End If
    'If musicNum = 18 then PlayMusic "Iron Maiden - Fear Of The Dark.mp3" End If
    'If musicNum = 19 then PlayMusic "Iron Maiden - Rime Of The Ancient Mariner.mp3" End If
    'If musicNum = 20 then PlayMusic "Iron Maiden - Murders In The Rue Morgue.mp3" End If
    'If musicNum = 21 then PlayMusic "Iron Maiden - Invaders.mp3" End If
    'If musicNum = 22 then PlayMusic "Iron Maiden - Comming Home.mp3" End If
    'If musicNum = 23 then PlayMusic "Iron Maiden - Gangland.mp3" End If
    'If musicNum = 24 then PlayMusic "Iron Maiden - The Number Of The Beast.mp3" End If

    musicNum = (musicNum + 1) mod 6
End Sub

'--------------------------------------------------------------------------------------------
'SOUNDS

  Sub TR_Drain_hit() PlaysoundAtVol "IM_ha", ActiveBall, 1: End Sub

  Sub TR_leftout_hit() PlaysoundAtVol "IM_nuuff", ActiveBall, 1: End Sub

  Sub TR_leftorb_hit() PlaysoundAtVol "IM_scream", ActiveBall, 1: End Sub

  Sub TR_rightout_hit() PlaysoundAtVol "IM_woe", ActiveBall, 1:End Sub

  Sub Target7_hit() PlaysoundAtVol "IM_ai", ActiveBall, 1:End Sub

  Sub Target8_hit() PlaysoundAtVol "IM_ai", ActiveBall, 1:End Sub

  Sub Target9_hit() PlaysoundAtVol "IM_nifnif", ActiveBall, 1:End Sub

  Sub Target10_hit() PlaysoundAtVol "IM_nifnif", ActiveBall, 1:End Sub


    playsound "IM_intro"

Sub Table1_Init()
	vpmInit Me
	With Controller
 		  .GameName = cGameName
          If Err Then MsgBox "Can't start Game " & cGameName & vbNewLine & Err.Description:Exit Sub
          .SplashInfoLine = "Firepower Williams 1980" & vbNewLine & "Created for VPX by WED21"
          .Games(cGameName).Settings.Value("sound") = 0
          .HandleKeyboard = 0
          .ShowTitle = 0
          .ShowFrame = 0
          .HandleMechanics = 0
		  .ShowDMDOnly = 1
		  .Hidden = 0
 		 .dip(0)=&h00  'Set to usa
          On Error Resume Next
         ' .Run GetPlayerHWnd
          If Err Then MsgBox Err.Description
          On Error Goto 0
      End With
		Controller.Run
PinMameTimer.enabled = 1
vpmMapLights AllLights


'----------------------------------------------
'Setup Ball Trough Object
Set trTrough = new cvpmTrough
trTrough.CreateEvents "trTrough", Array(Drain, BallRelease)
trTrough.balls = 3
trTrough.size = 3
trTrough.EntrySw = cOutHoleSW
trTrough.InitEntrySounds "DrainShort", "DrainShort", "DrainShort"
trTrough.InitExitSounds "BallRelease", "BallRelease"
trTrough.addsw 2,cLeftBallRampSW
trTrough.addsw 1, cCenterBallRampSW
trTrough.addsw 0, cRightBallRampSW
trTrough.initexit BallRelease, 90, 10
trTrough.StackExitBalls = 1
trTrough.MaxBallsPerKick = 1
trTrough.Reset

'----------------------------------------------

Dim DesktopMode: DesktopMode = Table1.ShowDT

If DesktopMode = True Then 'Show Desktop components
	RailLeft.visible=1
	RailRight.visible=1
	DisplayTimer7.enabled = True
	P1D1.visible = 1
	P1D2.visible = 1
	P1D3.visible = 1
	P1D4.visible = 1
	P1D5.visible = 1
	P1D6.visible = 1
	P1D7.visible = 1
	P2D1.visible = 1
	P2D2.visible = 1
	P2D3.visible = 1
	P2D4.visible = 1
	P2D5.visible = 1
	P2D6.visible = 1
	P2D7.visible = 1
	P3D1.visible = 1
	P3D2.visible = 1
	P3D3.visible = 1
	P3D4.visible = 1
	P3D5.visible = 1
	P3D6.visible = 1
	P3D7.visible = 1
	P4D1.visible = 1
	P4D2.visible = 1
	P4D3.visible = 1
	P4D4.visible = 1
	P4D5.visible = 1
	P4D6.visible = 1
	P4D7.visible = 1
	BaD1.visible = 1
	BaD2.visible = 1
	CrD1.visible = 1
	CrD2.visible = 1

Else
	RailLeft.visible=0
	RailRight.visible=0
	DisplayTimer7.enabled = False
	P1D1.visible = 0
	P1D2.visible = 0
	P1D3.visible = 0
	P1D4.visible = 0
	P1D5.visible = 0
	P1D6.visible = 0
	P1D7.visible = 0
	P2D1.visible = 0
	P2D2.visible = 0
	P2D3.visible = 0
	P2D4.visible = 0
	P2D5.visible = 0
	P2D6.visible = 0
	P2D7.visible = 0
	P3D1.visible = 0
	P3D2.visible = 0
	P3D3.visible = 0
	P3D4.visible = 0
	P3D5.visible = 0
	P3D6.visible = 0
	P3D7.visible = 0
	P4D1.visible = 0
	P4D2.visible = 0
	P4D3.visible = 0
	P4D4.visible = 0
	P4D5.visible = 0
	P4D6.visible = 0
	P4D7.visible = 0
	BaD1.visible = 0
	BaD2.visible = 0
	CrD1.visible = 0
	CrD2.visible = 0
	BGR1up.visible=0
	BGR2up.visible=0
	BGR3up.visible=0
	BGR4up.visible=0
	BGR1cp.visible=0
	BGR2cp.visible=0
	BGR3cp.visible=0
	BGR4cp.visible=0
	BGRtilt.visible=0
	BGRgo.visible=0
	BGRsa.visible=0
	BGRmatch.visible=0
	BGRbip.visible=0
	BGRhs.visible=0
end If

end Sub

Sub TargetSetup_timer()

If DTMod= 0 Then
DTarget1.visible=0
DTarget1.collidable=0
DTarget2.visible=0
DTarget2.collidable=0
DTarget3.visible=0
DTarget3.collidable=0
DTarget4.visible=0
DTarget4.collidable=0
DTarget5.visible=0
DTarget5.collidable=0
DTarget6.visible=0
DTarget6.collidable=0
Else
Target1.visible=0
Target1.collidable=0
Target2.visible=0
Target2.collidable=0
Target3.visible=0
Target3.collidable=0
Target4.visible=0
Target4.collidable=0
Target5.visible=0
Target5.collidable=0
Target6.visible=0
Target6.collidable=0
LTCheck.enabled=True
RTCheck.enabled=True
TopPOWERTarget.image="FirePowerTarget"
MiddlePOWERTarget.image="FirePowerTarget"
BottomPOWERTarget.image="FirePowerTarget"
End If
TargetSetup.Enabled=0
End Sub

'-----------------------------------------------------------------
' Setup Switch Subs

Sub LeftOutsideRollover_hit()
	vpmtimer.pulsesw cLeftOutsideRolloverSW
	if LShieldOn.state = lightstateon Then
		KBallSaveKicker.enabled = 1
	Else
		KballsaveKicker.enabled = 0
	end if
End Sub


Sub LeftInsideRollover_hit:vpmtimer.pulsesw cLeftInsideRolloverSW:Playsound "IM_flop":End Sub
Sub LeftInsideRollover_unhit:Controller.Switch(cLeftInsideRolloverSW)=0:End Sub



Sub UpperMiddleLeftStandup_hit:vpmtimer.pulsesw cUpperMiddleLeftStandupSW:Playsound "rubber_hit_3", 0, 1, -0.09, 0.09:End Sub
Sub UpperMiddleLeftStandup_unhit:Controller.Switch(cUpperMiddleLeftStandupSW)=0:End Sub

Sub Spinner_Spin:vpmTimer.PulseSw cSpinnerSW:PlaySoundAtVol "fx_spinner", Spinner, 1:End Sub

Sub TopLeftStandup_hit:vpmtimer.pulsesw cTopLeftStandupSW:Playsound "rubber_hit_3", 0, 1, -0.09, 0.09:End Sub
Sub TopLeftStandup_unhit:Controller.Switch(cTopLeftStandupSW)=0:End Sub


'----------------
'TargetSubs
Sub Target1_hit()
	vpmTimer.PulseSw cTarget1SW
	PlaysoundAtVol "Target", ActiveBall, 1
End Sub


Sub Target2_hit()
	vpmTimer.PulseSw cTarget2SW
	PlaysoundAtVol "Target", ActiveBall, 1
End Sub

Sub Target3_hit()
	vpmTimer.PulseSw cTarget3SW
	PlaysoundAtVol "Target", ActiveBall, 1
End Sub


Sub Target4_hit()
	vpmTimer.PulseSw cTarget4SW
	PlaysoundAtVol "Target", ActiveBall, 1
End Sub


Sub Target5_hit()
	vpmTimer.PulseSw cTarget5SW
	PlaysoundAtVol "Target", ActiveBall, 1
End Sub

Sub Target6_hit()
	vpmTimer.PulseSw cTarget6SW
	PlaysoundAtVol "Target", ActiveBall, 1
End Sub

Sub DTarget1_hit()
	vpmTimer.PulseSw cTarget1SW
	PlaysoundAtVol "Target", ActiveBall, 1
    Playsound "IM_E"
End Sub


Sub DTarget2_hit()
	vpmTimer.PulseSw cTarget2SW
	PlaysoundAtVol "Target", ActiveBall, 1
    Playsound "IM_D"
End Sub

Sub DTarget3_hit()
	vpmTimer.PulseSw cTarget3SW
	PlaysoundAtVol "Target", ActiveBall, 1
    Playsound "IM_W"
End Sub


Sub DTarget4_hit()
	vpmTimer.PulseSw cTarget4SW
	PlaysoundAtVol "Target", ActiveBall, 1
    Playsound "IM_A"
End Sub


Sub DTarget5_hit()
	vpmTimer.PulseSw cTarget5SW
	PlaysoundAtVol "Target", ActiveBall, 1
    Playsound "IM_R"
End Sub

Sub DTarget6_hit()
	vpmTimer.PulseSw cTarget6SW
	PlaysoundAtVol "Target", ActiveBall, 1
    Playsound "IM_D"
End Sub


Sub TopPowerTarget_hit()
	vpmTimer.PulseSw cTopPOWERTargetSW
	PlaysoundAtVol "Target", ActiveBall, 1
    Playsound "IM_maiden"
End Sub


Sub MIddlePOWERTarget_hit()
	vpmTimer.PulseSw cMiddlePOWERTargetSW
	PlaysoundAtVol "Target", ActiveBall, 1
    Playsound "IM_maiden"
End Sub


Sub BottomPOWERTarget_hit()
	vpmTimer.PulseSw cBottomPOWERTargetSW
	PlaysoundAtVol "Target", ActiveBall, 1
    Playsound "IM_maiden"
End Sub


Sub TopCenterTarget_hit()
	vpmTimer.PulseSw  cTopCenterTargetSW
	PlaysoundAtVol "Target", ActiveBall, 1
    Playsound "IM_solo"
End Sub

Sub Bumper4_hit:vpmTimer.PulseSw cBottomLeftBumperSW:PlaysoundAtVol "Fx_Bumper1", ActiveBall, 1:PlaysoundAtVol "IM_bumpers", ActiveBall, 1:End Sub

Sub Bumper1_hit:vpmTimer.PulseSw cTopLeftBumperSW:PlaysoundAtVol "Fx_Bumper2", ActiveBall, 1:PlaysoundAtVol "IM_bumpers", ActiveBall, 1:End Sub

Sub Bumper2_hit:vpmTimer.PulseSw cTopRightBumperSW:PlaysoundAtVol "Fx_Bumper3", ActiveBall, 1:PlaysoundAtVol "IM_bumpers", ActiveBall, 1:End Sub

Sub Bumper3_hit:vpmTimer.PulseSw cBottomRightBumperSW:PlaysoundAtVol "Fx_Bumper4", ActiveBall, 1:PlaysoundAtVol "IM_bumpers", ActiveBall, 1:End Sub




Sub FRollover_hit:vpmtimer.pulsesw cFRolloverSW:PlaysoundAtVol "IM_I", ActiveBall, 1:End Sub
Sub FRollover_unhit:Controller.Switch(cFRolloverSW)=0:End Sub

Sub IRollover_hit:vpmtimer.pulsesw cIRolloverSW:PlaysoundAtVol "IM_R", ActiveBall, 1:End Sub
Sub IRollover_unhit:Controller.Switch(cIRolloverSW)=0:End Sub

Sub RRollover_hit:vpmtimer.pulsesw cRRolloverSW:PlaysoundAtVol "IM_O", ActiveBall, 1:End Sub
Sub RRollover_unhit:Controller.Switch(cRRolloverSW)=0:End Sub

Sub ERollover_hit:vpmtimer.pulsesw cERolloverSW:PlaysoundAtVol "IM_N", ActiveBall, 1:End Sub
Sub ERollover_unhit:Controller.Switch(cERolloverSW)=0:End Sub


Sub LowerTopRightStandup_hit:vpmtimer.pulsesw cLowerTopRightStnadupSW:PlaysoundAtVol "rubber_hit_3", ActiveBall, 1 :End Sub
Sub LowerTopRightStandup_unhit:Controller.Switch(cLowerTopRightStnadupSW)=0:End Sub

Sub MiddleRightStandup_hit:vpmtimer.pulsesw cMIddleRightStandupSW:PlaysoundAtVol "rubber_hit_3", ActiveBall, 1 :End Sub
Sub MiddleRightStandup_unhit:Controller.Switch(cMIddleRightStandupSW)=0:End Sub


Sub RightInsideRollover_hit:vpmtimer.pulsesw cRightInsideRolloverSw:PlaysoundAtVol "IM_flop", ActiveBall, 1:End Sub
Sub RightInsideRollover_unhit:Controller.Switch(cRightInsideRolloverSW)=0:End Sub

Sub RightOutsideRollover_hit:vpmtimer.pulsesw cRightOutsideRolloverSW:End Sub
Sub RightOutsideRollover_unhit:Controller.Switch(cRightOutsideRolloverSW)=0:End Sub

Sub BallShooter_hit:Controller.Switch(cBallShooterSW)=1:End Sub
Sub BallShooter_unhit:Controller.Switch(cBallShooterSW)=0:End Sub

Sub LowerRightStandup_hit:vpmtimer.pulsesw cLowerRightStandupSW:PlaysoundAtVol "rubber_hit_3", ActiveBall, 1 :End Sub
Sub LowerRightStandup_unhit:Controller.Switch(cLowerRightStandupSW)=0:End Sub


Sub CenterMIddleLeftStandup_hit:vpmtimer.pulsesw cCenterMiddleLeftStandupSW:PlaysoundAtVol "rubber_hit_3", ActiveBall, 1 :End Sub
Sub CenterMIddleLeftStandup_unhit:Controller.Switch(cCenterMiddleLeftStandupSW)=0:End Sub

Sub LowerMIddleLeftStandup_hit:vpmtimer.pulsesw cLowerMiddleLeftStuandupSW:PlaysoundAtVol "rubber_hit_3", ActiveBall, 1 :End Sub
Sub LowerMIddleLeftStandup_unhit:Controller.Switch(cLowerMiddleLeftStuandupSW)=0:End Sub

Sub LeftEjectRollover_hit:vpmtimer.pulsesw cLeftEjectRolloverSW:End Sub
Sub LeftEjectRollover_unhit:Controller.Switch(cLeftEjectRolloverSW)=0:End Sub

Sub RightEjectRollover_hit:vpmtimer.pulsesw cRightEjectRolloverSW:End Sub
Sub RightEjectRollover_unhit:Controller.Switch(cRightEjectRolloverSW)=0:End Sub



'------------------------------------------------------------------
'Handle Keyboard Inputs

Sub Table1_KeyDown(ByVal keycode)
	If keycode = PlungerKey Then
		Plunger.PullBack
		PlaySoundAtVol "plungerpull", Plunger, 1
	End If

	If keycode = 4 or keycode = 5 or keycode = 6 Then
		PlaysoundAtVol "fx_coin", Drain, 1
	end if
    vpmKeyDown(keycode)
End Sub

Sub Table1_KeyUp(ByVal keycode)
	vpmKeyUp(keycode)
	If keycode = PlungerKey Then
		Plunger.Fire
		PlaySoundAtVol "plunger", Plunger, 1
	End If

End Sub

'---------------------
'Backglass
Sub Backglass_Timer()
If BGL1up.State=1 then BGR1up.SetValue(1) End If
If BGL1up.State=0 then BGR1up.SetValue(0) End If
If BGL2up.State=1 then BGR2up.SetValue(1) End If
If BGL2up.State=0 then BGR2up.SetValue(0) End If
If BGL3up.State=1 then BGR3up.SetValue(1) End If
If BGL3up.State=0 then BGR3up.SetValue(0) End If
If BGL4up.State=1 then BGR4up.SetValue(1) End If
If BGL4up.State=0 then BGR4up.SetValue(0) End If
If BGL1cp.State=1 then BGR1cp.SetValue(1) End If
If BGL1cp.State=0 then BGR1cp.SetValue(0) End If
If BGL2cp.State=1 then BGR2cp.SetValue(1) End If
If BGL2cp.State=0 then BGR2cp.SetValue(0) End If
If BGL3cp.State=1 then BGR3cp.SetValue(1) End If
If BGL3cp.State=0 then BGR3cp.SetValue(0) End If
If BGL4cp.State=1 then BGR4cp.SetValue(1) End If
If BGL4cp.State=0 then BGR4cp.SetValue(0) End If
If BGLbip.State=1 then BGRbip.SetValue(1) End If
If BGLbip.State=0 then BGRbip.SetValue(0) End If
If BGLmatch.State=1 then BGRmatch.SetValue(1) End If
If BGLmatch.State=0 then BGRmatch.SetValue(0) End If
If BGLgo.State=1 then BGRgo.SetValue(1) End If
If BGLgo.State=0 then BGRgo.SetValue(0) End If
If BGLsa.State=1 then BGRsa.SetValue(1) End If
If BGLsa.State=0 then BGRsa.SetValue(0) End If
If BGLhs.State=1 then BGRhs.SetValue(1) End If
If BGLhs.State=0 then BGRhs.SetValue(0) End If
If BGLtilt.State=1 then BGRtilt.SetValue(1) End If
If BGLtilt.State=0 then BGRtilt.SetValue(0) End If

End Sub

'--------------------------------------------------------------
'*********BALLKICKER************

Sub KBallSaveKicker_Hit()

		Me.kick 0, 35
		PlaysoundAtVol "SlingshotLeft", ActiveBall, 1
    PlaysoundAtVol "IM_666", ActiveBall, 1

End Sub


'--------------------------------------------------------------
' Play sounds when gates are Hit
Sub Gate3_Hit()
	PlaysoundAtVol "GateWire", ActiveBall, Vol(ActiveBall)
End Sub

Sub Gate1_Hit()
	PlaysoundAtVol "GateWire",ActiveBall, Vol(ActiveBall)
End Sub

Sub Gate2_Hit()
	PlaysoundAtVol "GateWire",ActiveBall, Vol(ActiveBall)
End Sub

Sub BallReleaseGate_Hit()
	PlaysoundAtVol "GateWire",Activeball, Vol(ActiveBall)
End Sub




 '=========================================================
'                    LED Handling
'=========================================================
'Modified version of Scapino's LED code for Fathom
'and borrowed from Uncle Willy's VP9 Firepower
'
Dim SixDigitOutput(32)
Dim SevenDigitOutput(32)
Dim DisplayPatterns(11)
Dim DigStorage(32)


'Binary/Hex Pattern Recognition Array
DisplayPatterns(0) = 0		'0000000 Blank
DisplayPatterns(1) = 63		'0111111 zero
DisplayPatterns(2) = 6		'0000110 one
DisplayPatterns(3) = 91		'1011011 two
DisplayPatterns(4) = 79		'1001111 three
DisplayPatterns(5) = 102	'1100110 four
DisplayPatterns(6) = 109	'1101101 five
DisplayPatterns(7) = 125	'1111101 six
DisplayPatterns(8) = 7		'0000111 seven
DisplayPatterns(9) = 127	'1111111 eight
DisplayPatterns(10)= 111	'1101111 nine

'Assign 7-digit output to reels
Set SevenDigitOutput(0)  = P1D7
Set SevenDigitOutput(1)  = P1D6
Set SevenDigitOutput(2)  = P1D5
Set SevenDigitOutput(3)  = P1D4
Set SevenDigitOutput(4)  = P1D3
Set SevenDigitOutput(5)  = P1D2
Set SevenDigitOutput(6)  = P1D1

Set SevenDigitOutput(7)  = P2D7
Set SevenDigitOutput(8)  = P2D6
Set SevenDigitOutput(9)  = P2D5
Set SevenDigitOutput(10) = P2D4
Set SevenDigitOutput(11) = P2D3
Set SevenDigitOutput(12) = P2D2
Set SevenDigitOutput(13) = P2D1

Set SevenDigitOutput(14) = P3D7
Set SevenDigitOutput(15) = P3D6
Set SevenDigitOutput(16) = P3D5
Set SevenDigitOutput(17) = P3D4
Set SevenDigitOutput(18) = P3D3
Set SevenDigitOutput(19) = P3D2
Set SevenDigitOutput(20) = P3D1

Set SevenDigitOutput(21) = P4D7
Set SevenDigitOutput(22) = P4D6
Set SevenDigitOutput(23) = P4D5
Set SevenDigitOutput(24) = P4D4
Set SevenDigitOutput(25) = P4D3
Set SevenDigitOutput(26) = P4D2
Set SevenDigitOutput(27) = P4D1

Set SevenDigitOutput(28) = CrD2
Set SevenDigitOutput(29) = CrD1
Set SevenDigitOutput(30) = BaD2
Set SevenDigitOutput(31) = BaD1

Sub DisplayTimer7_Timer ' 7-Digit output
	On Error Resume Next
	Dim ChgLED,ii,chg,stat,obj,TempCount,temptext,adj

	ChgLED = Controller.ChangedLEDs(&HFF, &HFFFF) 'hex of binary (display 111111, or first 6 digits)

	If Not IsEmpty(ChgLED) Then
		For ii = 0 To UBound(ChgLED)
			chg = chgLED(ii, 1) : stat = chgLED(ii, 2)
			For TempCount = 0 to 10
				If stat = DisplayPatterns(TempCount) then
					If LedStatus = 2 Then SevenDigitOutput(chgLED(ii, 0)).SetValue(TempCount)
					DigStorage(chgLED(ii, 0)) = TempCount
				End If
				If stat = (DisplayPatterns(TempCount) + 128) then
					If LedStatus = 2 Then SevenDigitOutput(chgLED(ii, 0)).SetValue(TempCount)
					DigStorage(chgLED(ii, 0)) = TempCount
				End If
			Next
		Next
	End IF
End Sub


'*DOF method for non rom controller tables by Arngrim****************
'*******Use DOF 1**, 1 to activate a ledwiz output*******************
'*******Use DOF 1**, 0 to deactivate a ledwiz output*****************
'*******Use DOF 1**, 2 to pulse a ledwiz output**********************
'Sub DOF(dofevent, dofstate)
'	If B2SOn=True Then
'		If dofstate = 2 Then
'			Controller.B2SSetData dofevent, 1:Controller.B2SSetData dofevent, 0
'		Else
'			Controller.B2SSetData dofevent, dofstate
'		End If
'	End If
'End Sub
'********************************************************************

' *******************************************************************************************************
' Positional Sound Playback Functions by DJRobX and Rothbauerw
' PlaySound sound, 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 1, AudioFade(ActiveBall)
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

Sub PlaySoundAtVol(sound, tableobj, Volume)
  PlaySound sound, 1, Volume, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

'Set all as per ball position & speed, but Vol Multiplier may be used eg; PlaySoundAtBallVol "sound",3

Sub PlaySoundAtBallVol(sound, VolMult)
  PlaySound sound, 0, Vol(ActiveBall) * VolMult, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 1, AudioFade(ActiveBall)
End Sub

'Set position as bumperX and Vol manually.

Sub PlaySoundAtBumperVol(sound, tableobj, Vol)
  PlaySound sound, 1, Vol, AudioPan(tableobj), 0,0,1, 1, AudioFade(tableobj)
End Sub

Sub PlaySoundAtBOTBallZ(sound, BOT)
		PlaySound sound, 0, ABS(BOT.velz)/17, Pan(BOT), 0, Pitch(BOT), 1, 0, AudioFade(BOT)
End Sub

'*********************************************************************
'                     Supporting Ball & Sound Functions
'*********************************************************************

Function RndNum(min, max)
    RndNum = Int(Rnd() * (max-min + 1) ) + min ' Sets a random number between min and max
End Function

Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "table1" is the name of the table
  Dim tmp
  On Error Resume Next
  tmp = tableobj.y * 2 / table1.height-1
  If tmp > 0 Then
    AudioFade = Csng(tmp ^10)
  Else
    AudioFade = Csng(-((- tmp) ^10) )
  End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "table1" is the name of the table
  Dim tmp
  On Error Resume Next
  tmp = tableobj.x * 2 / table1.width-1
  If tmp > 0 Then
    AudioPan = Csng(tmp ^10)
  Else
    AudioPan = Csng(-((- tmp) ^10) )
  End If
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "table1" is the name of the table
  Dim tmp
  On Error Resume Next
  tmp = ball.x * 2 / table1.width-1
  If tmp > 0 Then
    Pan = Csng(tmp ^10)
  Else
    Pan = Csng(-((- tmp) ^10) )
  End If
End Function

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
  Vol = Csng(BallVel(ball) ^2 / VolDiv)
End Function

Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
  Pitch = BallVel(ball) * 20
End Function

Function BallVel(ball) 'Calculates the ball speed
  BallVel = INT(SQR((ball.VelX ^2) + (ball.VelY ^2) ) )
End Function

Function BallVelZ(ball) 'Calculates the ball speed in the -Z
    BallVelZ = INT((ball.VelZ) * -1 )
End Function

Function VolZ(ball) ' Calculates the Volume of the sound based on the ball speed in the Z
    VolZ = Csng(BallVelZ(ball) ^2 / 200)*1.2
End Function

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

'*****************************************
'      JP's VP10 Rolling Sounds
'*****************************************

Const tnob = 4 ' total number of balls
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
      If BallVel(BOT(b) ) > 1 Then
        rolling(b) = True
        if BOT(b).z < 30 Then ' Ball on playfield
          PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) ), AudioPan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0, AudioFade(BOT(b) )
        Else ' Ball on raised ramp
          PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) )*.5, AudioPan(BOT(b) ), 0, Pitch(BOT(b) )+50000, 1, 0, AudioFade(BOT(b) )
        End If
      Else
        If rolling(b) = True Then
          StopSound("fx_ballrolling" & b)
          rolling(b) = False
        End If
      End If
		If BOT(b).VelZ < -1 and BOT(b).z < 55 and BOT(b).z > 27 Then 'height adjust for ball drop sounds
			PlaySoundAtBOTBallZ "fx_ball_drop" & b, BOT(b)
			'debug.print BOT(b).velz
		End If
    Next
End Sub

'**********************
' Ball Collision Sound
'**********************

Sub OnBallBallCollision(ball1, ball2, velocity)
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / (VolDiv/VolCol), AudioPan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub

'*************SUPPORTING SOUNDS************************

Sub Pins_Hit (idx)
	PlaySound "pinhit_low", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Targets_Hit (idx)
	PlaySound "target", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Switches_Hit (idx)
	PlaySound "metalhit_thin", 0, 0.75, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Metals_Medium_Hit (idx)
	PlaySound "metalhit_medium", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Metals2_Hit (idx)
	PlaySound "metalhit2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Rubbers_Hit(idx)
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 20 then
		PlaySound "fx_rubber2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End if
	If finalspeed >= 6 AND finalspeed <= 20 then
 		RandomSoundRubber()
 	End If
End Sub

Sub Posts_Hit(idx)
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 16 then
		PlaySound "fx_rubber2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End if
	If finalspeed >= 6 AND finalspeed <= 16 then
 		RandomSoundRubber()
 	End If
End Sub

Sub RandomSoundRubber()
	Select Case Int(Rnd*3)+1
		Case 1 : PlaySound "rubber_hit_1", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 2 : PlaySound "rubber_hit_2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 3 : PlaySound "rubber_hit_3", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
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
		Case 1 : PlaySound "flip_hit_1", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 2 : PlaySound "flip_hit_2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 3 : PlaySound "flip_hit_3", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End Select
End Sub


'*********** BALL SHADOW *********************************
Dim BallShadow
BallShadow = Array (BallShadow1,BallShadow2,BallShadow3)

Sub BallShadowUpdate_timer()
    Dim BOT, b
    BOT = GetBalls
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

Sub Table1_Exit():Controller.Games(cGameName).Settings.Value("sound") = 1:Controller.Stop:End Sub

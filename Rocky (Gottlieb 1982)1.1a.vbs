Option Explicit
Randomize

Const cGameName = "rocky"

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "Can't open controller.vbs"
On Error Goto 0

LoadVPM "01000100", "sys80.vbs", 2.31

'**********************************************************
'********       OPTIONS     *******************************
'**********************************************************

Dim BallShadows: Ballshadows=1          '******************set to 1 to turn on Ball shadows
Dim FlipperShadows: FlipperShadows=1  '***********set to 1 to turn on Flipper shadows

'************************************************
'************************************************
'************************************************
'************************************************
'************************************************

Const UseSolenoids = True
Const UseLamps = True
Const UseSync = False
Const UseGI = False

' Thalamus 2019 July : Improved directional sounds
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


' Standard Sounds
Const SSolenoidOn = "fx_solenoid"
Const SSolenoidOff = "fx_solenoidoff"
Const SCoin = "fx_coin"

Dim bsTrough, bsSaucer, bsKicker, DTBankL, DTBankR, DTBankT, FastFlips

DisplayTimer.Enabled = true

'****Table Init****
  Sub Table1_Init
  	With Controller
        .GameName 								= cGameName
        .SplashInfoLine 						= "Rocky, Gottleib 1982"
        .HandleKeyboard = 0
         .ShowTitle = 0
         .ShowDMDOnly = 1
         .ShowFrame = 0
         .HandleMechanics = False
'		 .Hidden = HiddenValue
         On Error Resume Next
         .Run GetPlayerHWnd
         If Err Then MsgBox Err.Description
         On Error Goto 0
     End With
	 Controller.SolMask(0)=0
	 vpmTimer.AddTimer 2000,"Controller.SolMask(0)=&Hffffffff'" 'ignore all solenoids - then add the timer to renable all the solenoids after 2 s

    PinMAMETimer.Interval = PinMAMEInterval
    PinMAMETimer.Enabled = 1

	' Nudging
    vpmNudge.TiltSwitch = 57
     vpmNudge.Sensitivity = 3
     vpmNudge.TiltObj = Array(leftslingshot, lleftslingshot,lleftslingshot1,rightslingshot, rrightslingshot,bumper1, bumper2, bumper3, bumper4)


	'**Trough
          Set bsTrough = New cvpmBallStack
            With bsTrough
             .InitNoTrough BallRelease, 67, 90, 10
             .InitExitSnd "ballrelease", "Solenoid"
            End With

	'**Saucers
           Set bsSaucer = New cvpmBallStack  'Top Suacer
             With bsSaucer
			  .InitSaucer sw4,4, 230, 8

      		  .InitExitSnd "ballrelease", "Solenoid"
    		  .InitAddSnd "kicker_enter"

             End With

			Set bsKicker = New cvpmBallStack  'Top Suacer
             With bsKicker
               .InitSaucer sw24,24, 0, 40
      		   .InitExitSnd "FX_Kicker", "Solenoid"
    		   .InitAddSnd "Target"
             End With

	'**DropTargets
   	Set DTBankL=New cvpmDropTarget  'Left
   	    DTBankL.InitDrop Array(sw2,sw12,sw22), Array(2,12,22)
        DTBankL.InitSnd SoundFX("fx_droptarget",DOFDropTargets),SoundFX("fx_DTReset",DOFDropTargets)

   	Set DTBankR=New cvpmDropTarget 'right
		DTBankR.InitDrop Array(sw1,sw11,sw21), Array(1,11,21)
        DTBankR.InitSnd SoundFX("fx_droptarget",DOFDropTargets),SoundFX("fx_DTReset",DOFDropTargets)


   	Set DTBankT=New cvpmDropTarget 'Blue Trap Targets
   	    DTBankT.InitDrop Array(sw00,sw10,sw20,sw50,sw60), Array(0,10,20,50,60)
        DTBankT.InitSnd SoundFX("fx_droptarget",DOFDropTargets),SoundFX("fx_DTReset",DOFDropTargets)


 Set FastFlips = new cFastFlips
    with FastFlips
        .CallBackL = "SolLflipper"  'Point these to flipper subs
        .CallBackR = "SolRflipper"  '...
        .TiltObjects = True 'Optional, if True calls vpmnudge.solgameon automatically. IF YOU GET A LINE 1 ERROR, DISABLE THIS! (or setup vpmNudge.TiltObj!)
    '   .DebugOn = False        'Debug, always-on flippers. Call FastFlips.DebugOn True or False in debugger to enable/disable.
    end with

  End Sub

'************************************************
' Solenoids
'************************************************
SolCallback(1) =   "bsSaucer.SolOut"
SolCallback(2) =   "DTBankTReset"
SolCallback(3) =   "SolKicker"
SolCallback(4) =   "vpmSolSound ""Bell"","
SolCallback(5) =   "DTBankRReset"
SolCallback(6) =   "DTBankLReset"
SolCallback(8) =   "vpmSolSound ""Knocker"","
SolCallback(9) = "solBallRelease"
SolCallback(10) =  "FastFlips.TiltSol"

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Dim timer

Sub SolLFlipper(Enabled)
     If Enabled Then
         PlaySoundAtVol SoundFX("fx_Flipperup",DOFContactors), LeftFlipper, VolFlip:LeftFlipper.RotateToEnd:LeftFlipper1.RotateToEnd:LeftFlipper2.RotateToEnd':vpmTimer.PulseSw 62 no idea
     Else
         PlaySoundAtVol SoundFX("fx_Flipperdown",DOFContactors), LeftFlipper, VolFlip:LeftFlipper.RotateToStart:LeftFlipper1.RotateToStart:LeftFlipper2.RotateToStart
     End If
  End Sub

Sub SolRFlipper(Enabled)
     If Enabled Then
         PlaySoundAtVol SoundFX("fx_Flipperup",DOFContactors), RightFlipper, VolFlip:RightFlipper.RotateToEnd:RightFlipper1.RotateToEnd:vpmTimer.PulseSw 61
     Else
         PlaySoundAtVol SoundFX("fx_Flipperdown",DOFContactors), RightFlipper, VolFlip:RightFlipper.RotateToStart:RightFlipper1.RotateToStart
     End If
End Sub

Sub solBallRelease(Enabled)
	If enabled then
	bstrough.ExitSol_On 'solexit ssolenoidon, ssolenoidon,enabled
	End if
  End Sub

  Sub SolKicker(Enabled)
	If enabled then
		bsKicker.ExitSol_On
	End If
  End Sub

Sub Table1_KeyDown(ByVal KeyCode)
    If KeyCode = LeftFlipperKey then FastFlips.FlipL True :  FastFlips.FlipUL True
    If KeyCode = RightFlipperKey then FastFlips.FlipR True :  FastFlips.FlipUR True
	If Keycode = LeftMagnaSave Then vpmTimer.PulseSw 66
    If KeyDownHandler(keycode) Then Exit Sub
    If keycode = PlungerKey Then Plunger.Pullback:playsoundAtVol"plungerpull", Plunger, 1
End Sub

Sub Table1_KeyUp(ByVal KeyCode)
    If KeyCode = LeftFlipperKey then FastFlips.FlipL False :  FastFlips.FlipUL False
    If KeyCode = RightFlipperKey then FastFlips.FlipR False :  FastFlips.FlipUR False
    If KeyUpHandler(keycode) Then Exit Sub
    If keycode = PlungerKey Then Plunger.Fire:PlaySoundAtVol"plunger", Plunger, 1
End Sub

Sub Drain_Hit():PlaySoundAtVol "drain", Drain, 1:bstrough.addball me:End Sub':vpmTimer.PulseSw 67:End Sub
Sub sw4_Hit():bsSaucer.AddBall Me:RightEMpos = 0:PkickarmR.RotZ = 2:RightEMTimer.Enabled = 1:End Sub
Sub sw24_Hit():bsKicker.AddBall Me:End Sub
Sub sw24_UnHit:bsKickerP.rotx = 10:bsKickerP.rotx = 20:bsKickerP.rotx = 10:bsKickerP.rotx = 0:End Sub

'*****Drop Lights Off
   dim xx
    For each xx in DTBankRlights: xx.state=0:Next
	For each xx in DTBankLlights: xx.state=0:Next
	For each xx in DTBankTlights: xx.state=0:Next

'DropTargets
 Sub sw2_Hit:DTBankL.Hit 1:sw2L1.state=1:End Sub
 Sub sw12_Hit:DTBankL.Hit 2:sw12L1.state=1:End Sub
 Sub sw22_Hit:DTBankL.Hit 3:sw22L1.state=1:End Sub

 Sub sw1_Hit:DTBankR.Hit 1:sw1L1.state=1:End Sub
 Sub sw11_Hit:DTBankR.Hit 2:sw11L1.state=1:End Sub
 Sub sw21_Hit:DTBankR.Hit 3:sw21L1.state=1:End Sub

 Sub sw00_Hit:DTBankT.Hit 1:sw00L1.state=1::End Sub
 Sub sw10_Hit:DTBankT.Hit 2:sw10L1.state=1::End Sub
 Sub sw20_Hit:DTBankT.Hit 3:sw20L1.state=1::End Sub
 Sub sw50_Hit:DTBankT.Hit 4:sw50L1.state=1::End Sub
 Sub sw60_Hit:DTBankT.Hit 5:sw60L1.state=1::End Sub

'DropTargets
   	Sub DTBankLReset(enabled)  'Left
   	  dim xx
		if enabled then
   		DTBankL.SolDropUp enabled
		For each xx in DTBankLLights: xx.state=0:Next
         end if
 End Sub

   	Sub DTBankRReset(enabled) 'right
   	  dim xx
		if enabled then
   		DTBankR.SolDropUp enabled
		For each xx in DTBankRLights: xx.state=0:Next
         end if
 End Sub

   	Sub DTBankTReset(enabled) 'Top Targets
   	  dim xx
		if enabled then
   		DTBankT.SolDropUp enabled
		For each xx in DTBankTLights: xx.state=0:Next
         end if
 End Sub

'Spinners
  Sub sw55_Spin():PlaySoundAtVol "fx_spinner", sw55, VolSpin:vpmtimer.pulsesw 55:End Sub

'Bumpers

Sub bumper1_Hit : vpmTimer.PulseSw 3 : playsoundAtVol SoundFX("fx_bumper1",DOFContactors),ActiveBall,VolBump: DOF 206, DOFPulse:End Sub
Sub bumper2_Hit : vpmTimer.PulseSw 13 : playsoundAtVol SoundFX("fx_bumper2",DOFContactors),ActiveBall,VolBump: DOF 209, DOFPulse:End Sub
Sub bumper3_Hit : vpmTimer.PulseSw 63 : playsoundAtVol SoundFX("fx_bumper3",DOFContactors),ActiveBall,VolBump: DOF 208, DOFPulse:End Sub
Sub bumper4_Hit : vpmTimer.PulseSw 23 : playsoundAtVol SoundFX("fx_bumper4",DOFContactors),ActiveBall,VolBump: DOF 207, DOFPulse:End Sub

'Wire Triggers
Sub sw5_Hit:Controller.Switch(5) = 1:PlaySoundAtVol "rollover", ActiveBall, 1:End Sub
Sub sw5_UnHit:Controller.Switch(5) = 0:End Sub
Sub sw15_Hit:Controller.Switch(15) = 1:PlaySoundAtVol "rollover", ActiveBall, 1:End Sub
Sub sw15_UnHit:Controller.Switch(15) = 0:End Sub
Sub sw25_Hit:Controller.Switch(25) = 1:PlaySoundAtVol "rollover", ActiveBall, 1:End Sub
Sub sw25_UnHit:Controller.Switch(25) = 0:End Sub
Sub sw14_Hit:Controller.Switch(14) = 1:PlaySoundAtVol "rollover", ActiveBall, 1:End Sub
Sub sw14_UnHit:Controller.Switch(14) = 0:End Sub
Sub sw51_Hit:Controller.Switch(51) = 1:PlaySoundAtVol "rollover", ActiveBall, 1:End Sub
Sub sw51_UnHit:Controller.Switch(51) = 0:End Sub
Sub sw52_Hit:Controller.Switch(52) = 1:PlaySoundAtVol "rollover", ActiveBall, 1:End Sub
Sub sw52_UnHit:Controller.Switch(52) = 0:End Sub
Sub sw54_Hit:Controller.Switch(54) = 1:PlaySoundAtVol "rollover", ActiveBall, 1:End Sub
Sub sw54_UnHit:Controller.Switch(54) = 0:End Sub
Sub sw64_Hit:Controller.Switch(64) = 1:PlaySoundAtVol "rollover", ActiveBall, 1:End Sub
Sub sw64_UnHit:Controller.Switch(64) = 0:End Sub
Sub center_Hit:PlaySoundAtVol "rollover", ActiveBall, 1:End Sub
Sub Trigger2_Hit:End Sub

Dim ccStep1
Sub Trigger2_unhit
    PlaySoundAtVol SoundFX("rollover",DOFContactors),ActiveBall,1
    DOF 201, DOFPulse
    bsKickerP.rotx = 0
    ccStep1 = 0
    Trigger2.TimerEnabled = 1
End Sub

Sub Trigger2_Timer
    Select Case ccStep1
		Case 2:bsKickerP.rotx = 25
        Case 3:bsKickerP.rotx = 10
        Case 4:bsKickerP.rotx = 0:Trigger2.TimerEnabled = 0
    End Select
    ccStep1 = ccStep1 + 1
End Sub

'Targets
 sub sw65_hit:vpmTimer.PulseSw 65:PlaySoundAtVol "Target", ActiveBall, 1:end sub

'Gate
  Sub SW6_Hit():PlaysoundAtVol "Gate", ActiveBall, 1:vpmtimer.pulsesw 6:End Sub

Sub SolKnocker(Enabled)
    If Enabled Then PlaySound SoundFX("Knocker",DOFKnocker)
End Sub

Sub table1_Paused:Controller.Pause = 1:End Sub

Sub table1_unPaused:Controller.Pause = 0:End Sub

Sub table1_Exit
    If b2son then controller.stop
End Sub

'GI uselamps workaround
dim GIlamps : set GIlamps = New GIcatcherobject
Class GIcatcherObject   'object that disguises itself as a light. (UseLamps workaround for System80 GI circuit)
    Public Property Let State(input)
        dim x
        if input = 1 then 'If GI switch is engaged, turn off GI.
            for each x in gi : x.state = 0 : next
        elseif input = 0 then
            for each x in gi : x.state = 1 : next
        end if
        'tb.text = "gitcatcher.state = " & input    'debug
    End Property
End Class

'****Lights***
set Lights(1) = GIlamps 'GI circuit
set lights(3)=Light3
set lights(4)=light4
set lights(5)=light5
set lights(6)=light6
set lights(7)=light7
set lights(8)=light8
lights(22)= Array(light22,Light22a)
lights(23)= Array(light23,Light23a)
set lights(24)=light24
set lights(25)=light25
set lights(26)=light26
set lights(27)=light27
lights(28)= Array(light28,Light28a)
lights(29)= Array(light29,Light29a)
lights(30)= Array(light30,Light30a)
lights(31)= Array(light31,Light31a)
set lights(32)=light32
set lights(33)=light33
set lights(34)=light34
set lights(35)=light35
set lights(36)=light36
set lights(37)=light37
set lights(38)=light38
set lights(39)=light39
set lights(40)=light40
set lights(41)=light41
set lights(42)=light42
set lights(43)=light43
set lights(44)=light44
set lights(45)=light45
set lights(46)=light46
set lights(47)=light47
set lights(51)=light51 'NO IDEA IF RIGHT
set lights(55)=Light55

'***** EM kicker animation

Dim RightEMPos

Sub Solout(enabled)
	If enabled Then
		Solout.ExitSol_On
		RightEMpos = 0
		PkickarmR.RotZ = 2
		RightEMTimer.Enabled = 1
	End If
End Sub

Sub RightEMTimer_Timer
    Select Case RightEMpos
        Case 1:PkickarmR.Rotz = 15
        Case 2:PkickarmR.Rotz = 15
        Case 3:PkickarmR.Rotz = 15
        Case 4:PkickarmR.Rotz = 8
        Case 5:PkickarmR.Rotz = 4
        Case 6:PkickarmR.Rotz = 2
        Case 7:PkickarmR.Rotz = 0:RightEMTimer.Enabled = 0
    End Select
    RightEMpos = RightEMpos + 1
End Sub

'**Rubber Point Walls
Sub sw16a_Hit:vpmtimer.pulsesw(16):End Sub
Sub sw16b_Hit:vpmtimer.pulsesw(16):End Sub
Sub sw16c_Hit:vpmtimer.pulsesw(16):End Sub

'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep, RRStep, LLStep, LLStep1

Sub RightSlingShot_Slingshot
    PlaySoundAtVol SoundFX("fx_slingshot",DOFContactors), sling1, 1
    DOF 202, DOFPulse
    vpmtimer.PulseSw(16)
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.rotx = 20
    RStep = 0
    RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.rotx = 10
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.rotx = 0:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
    PlaySoundAtVol SoundFX("fx_slingshot",DOFContactors), sling2, 1
    DOF 201, DOFPulse
    vpmtimer.pulsesw(16)
    LSling.Visible = 0
    LSling1.Visible = 1
    sling2.rotx = 20
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.rotx = 10
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.rotx = 0:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

Sub RRightSlingShot_Slingshot
    PlaySoundAtVol SoundFX("fx_slingshot",DOFContactors), sling3, 1
    DOF 202, DOFPulse
    vpmtimer.PulseSw(16)
    RRSling.Visible = 0
    RRSling1.Visible = 1
    Sling3.rotx = 20
    RRStep = 0
    RRightSlingShot.TimerEnabled = 1
End Sub

Sub RRightSlingShot_Timer
    Select Case RRStep
        Case 3:RRSLing1.Visible = 0:RRSLing2.Visible = 1:Sling3.rotx = 10
        Case 4:RRSLing2.Visible = 0:RRSLing.Visible = 1:Sling3.rotx = 0:RRightSlingShot.TimerEnabled = 0
    End Select
    RRStep = RRStep + 1
End Sub

Sub LLeftSlingShot_Slingshot
    PlaySoundAtVol SoundFX("fx_slingshot",DOFContactors), sling4, 1
    DOF 201, DOFPulse
    vpmtimer.pulsesw(16)
    LLSling.Visible = 0
    LLSling1.Visible = 1
    sling4.rotx = 20
    LLStep = 0
    LLeftSlingShot.TimerEnabled = 1
End Sub

Sub LLeftSlingShot_Timer
    Select Case LLStep
        Case 3:LLSLing1.Visible = 0:LLSLing2.Visible = 1:sling4.rotx = 10
        Case 4:LLSLing2.Visible = 0:LLSLing.Visible = 1:sling4.rotx = 0:LLeftSlingShot.TimerEnabled = 0
    End Select
    LLStep = LLStep + 1
End Sub

Sub LLeftSlingShot1_Slingshot
    PlaySoundAtVol SoundFX("fx_slingshot",DOFContactors), sling5, 1
    DOF 201, DOFPulse
    vpmtimer.pulsesw(16)
    LLSling4.Visible = 0
    LLSling3.Visible = 1
    sling5.rotx = 20
    LLStep1 = 0
    LLeftSlingShot1.TimerEnabled = 1
End Sub

Sub LLeftSlingShot1_Timer
    Select Case LLStep1
        Case 3:LLSLing3.Visible = 0:LLSLing5.Visible = 1:sling5.rotx = 10
        Case 4:LLSLing5.Visible = 0:LLSLing4.Visible = 1:sling5.rotx = 0:LLeftSlingShot1.TimerEnabled = 0
    End Select
    LLStep1 = LLStep1 + 1
End Sub

Dim Digits(32)

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'           LL      EEEEEE  DDDD        ,,   SSSSS
'           LL      EE      DD  DD      ,,  SS
'           LL      EE      DD   DD      ,   SS
'           LL      EEEE    DD   DD            SS
'           LL      EE      DD  DD              SS
'           LLLLLL  EEEEEE  DDDD            SSSSS
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
'       7 Digit Array
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

Dim LED7(35)
LED7(0)=Array()'d261,d262,d263,d264,d265,d266,d267,LXM,d268)
LED7(1)=Array()'d271,d272,d273,d274,d275,d276,d277,LXM,d278)
LED7(2)=Array()'d241,d242,d243,d244,d245,d246,d247,LXM,d248)
LED7(3)=Array()'d251,d252,d253,d254,d255,d256,d257,LXM,d258)
LED7(4)=Array()'d261,d262,d263,d264,d265,d266,d267,LXM,d268)
LED7(5)=Array()'d271,d272,d273,d274,d275,d276,d277,LXM,d278)
LED7(6)=Array()'d241,d242,d243,d244,d245,d246,d247,LXM,d248)
LED7(7)=Array()'d251,d252,d253,d254,d255,d256,d257,LXM,d258)
LED7(8)=Array()'d261,d262,d263,d264,d265,d266,d267,LXM,d268)
LED7(9)=Array()'d271,d272,d273,d274,d275,d276,d277,LXM,d278)
LED7(10)=Array()'d241,d242,d243,d244,d245,d246,d247,LXM,d248)
LED7(11)=Array()'d251,d252,d253,d254,d255,d256,d257,LXM,d258)
LED7(12)=Array()'d261,d262,d263,d264,d265,d266,d267,LXM,d268)
LED7(13)=Array()'d271,d272,d273,d274,d275,d276,d277,LXM,d278)
LED7(14)=Array()'d241,d242,d243,d244,d245,d246,d247,LXM,d248)
LED7(15)=Array()'d251,d252,d253,d254,d255,d256,d257,LXM,d258)
LED7(16)=Array()'d261,d262,d263,d264,d265,d266,d267,LXM,d268)
LED7(17)=Array()'d271,d272,d273,d274,d275,d276,d277,LXM,d278)
LED7(18)=Array()'d241,d242,d243,d244,d245,d246,d247,LXM,d248)
LED7(19)=Array()'d251,d252,d253,d254,d255,d256,d257,LXM,d258)
LED7(20)=Array()'d261,d262,d263,d264,d265,d266,d267,LXM,d268)
LED7(21)=Array()'d271,d272,d273,d274,d275,d276,d277,LXM,d278)
LED7(22)=Array()'d241,d242,d243,d244,d245,d246,d247,LXM,d248)
LED7(23)=Array()'d251,d252,d253,d254,d255,d256,d257,LXM,d258)


LED7(24)=Array()'d261,d262,d263,d264,d265,d266,d267,LXM,d268)
LED7(25)=Array()'d271,d272,d273,d274,d275,d276,d277,LXM,d278)
LED7(26)=Array()'d241,d242,d243,d244,d245,d246,d247,LXM,d248)
LED7(27)=Array()'d251,d252,d253,d254,d255,d256,d257,LXM,d258)

LED7(28)=Array()'d261,d262,d263,d264,d265,d266,d267,LXM,d268)
LED7(29)=Array()'d271,d272,d273,d274,d275,d276,d277,LXM,d278)
LED7(30)=Array()'d241,d242,d243,d244,d245,d246,d247,LXM,d248)
LED7(31)=Array()'d251,d252,d253,d254,d255,d256,d257,LXM,d258)
LED7(32)=Array(d241,d242,d243,d244,d245,d246,d247,LXM,d248)
LED7(33)=Array(d251,d252,d253,d254,d255,d256,d257,LXM,d258)
LED7(34)=Array(d261,d262,d263,d264,d265,d266,d267,LXM,d268)
LED7(35)=Array(d271,d272,d273,d274,d275,d276,d277,LXM,d278)

Sub DisplayTimer_Timer
    Dim ChgLED, II, Num, Chg, Stat, Obj
    ChgLED = Controller.ChangedLEDs(&HFF, &HFFFF)
    If Not IsEmpty(ChgLED) Then
        For II = 0 To UBound(ChgLED)
            Num = ChgLED(II, 0):Chg = ChgLED(II, 1):Stat = ChgLED(II, 2)
            If Num > 23 Then
                For Each Obj In LED7(Num)
                    If Chg And 1 Then Obj.State = Stat And 1

                    Chg = Chg \ 2:Stat = Stat \ 2
                Next
            End If
        Next
    End If
End Sub

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
    '***Ball Drop Sounds***

		If BOT(b).VelZ < -1 and BOT(b).z < 55 and BOT(b).z > 27 Then 'height adjust for ball drop sounds
			PlaySound "fx_ball_drop" & b, 0, ABS(BOT(b).velz)/17, AudioPan(BOT(b)), 0, Pitch(BOT(b)), 1, 0, AudioFade(BOT(b))
		End If

    Next
End Sub

'*****************************************
'	ninuzzu's	FLIPPER SHADOWS
'*****************************************

sub FlipperTimer_Timer()
	FlipperLSh.RotZ = LeftFlipper.currentangle
	FlipperLSh1.RotZ = LeftFlipper.currentangle
	FlipperRSh.RotZ = RightFlipper.currentangle
	FlipperRSh1.RotZ = RightFlipper.currentangle
	FlipperLSh2.RotZ = LeftFlipper2.currentangle
	LFlip.rotz = LeftFlipper.currentangle
	RFlip.rotz = RightFlipper.currentangle
	LflipRubber.rotz = LeftFlipper.currentangle
	RFlipRubber.rotz = RightFlipper.currentangle
	LFlip1.rotz = LeftFlipper1.currentangle
	RFlip1.rotz = RightFlipper1.currentangle
	LflipRubber1.rotz = LeftFlipper1.currentangle
	RFlipRubber1.rotz = RightFlipper1.currentangle
	LFlip2.rotz = LeftFlipper2.currentangle
	LflipRubber2.rotz = LeftFlipper2.currentangle
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

' The sound is played using the VOL, AUDIOPAN, AUDIOFADE and PITCH functions, so the volume and pitch of the sound
' will change according to the ball speed, and the AUDIOPAN & AUDIOFADE functions will change the stereo position
' according to the position of the ball on the table.


'**************************************
' Explanation of the collision routine
'**************************************

' The collision is built in VP.
' You only need to add a Sub OnBallBallCollision(ball1, ball2, velocity) and when two balls collide they
' will call this routine. What you add in the sub is up to you. As an example is a simple Playsound with volume and paning
' depending of the speed of the collision.


'**********************
' Ball Collision Sound
'**********************

Sub OnBallBallCollision(ball1, ball2, velocity)
	PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 2000, AudioPan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub

'**********************
' Object Hit Sounds
'**********************
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
 	If finalspeed > 5 then
		PlaySound "fx_rubber2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End if
	If finalspeed >= 2 AND finalspeed <= 5 then
 		RandomSoundRubber()
 	End If
End Sub

Sub Posts_Hit(idx)
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 16 then
		PlaySound "rubber_hit_2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
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


Sub LeftFlipper2_Collide(parm)
 	RandomSoundFlipper()
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

Sub RightFlipper1_Collide(parm)
 	RandomSoundFlipper()
End Sub

Sub RandomSoundFlipper()
	Select Case Int(Rnd*3)+1
		Case 1 : PlaySoundAtBall "flip_hit_1"
		Case 2 : PlaySoundAtBall "flip_hit_2"
		Case 3 : PlaySoundAtBall "flip_hit_3"
	End Select
End Sub

'cFastFlips by nFozzy
'Bypasses pinmame callback for faster and more responsive flippers
'Version 1.1 beta2 (More proper behaviour, extra safety against script errors)
'*************************************************
Function NullFunction(aEnabled):End Function    '1 argument null function placeholder
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
    Public Sub InitDelay(aName, aDelay) : Name = aName : delay = aDelay : End Sub   'Create Delay
    'Automatically decouple flipper solcallback script lines (only if both are pointing to the same sub) thanks gtxjoe
    Private Sub Decouple(aSolType, aInput)  : If StrComp(SolCallback(aSolType),aInput,1) = 0 then SolCallback(aSolType) = Empty End If : End Sub

    'call callbacks
    Public Sub FlipL(aEnabled)
        FlipState(0) = aEnabled 'track flipper button states: the game-on sol flips immediately if the button is held down (1.1)
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

    Public Sub TiltSol(aEnabled)    'Handle solenoid / Delay (if delayinit)
        If delay > 0 and not aEnabled then  'handle delay
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


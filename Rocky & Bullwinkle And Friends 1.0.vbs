'                    **************\/***************
'
'                   **************************************
' ******************** Rocky & Bullwinkle And Friends **********************
' ********************    DATA EAST 1992 FS Ver. 0.96  **********************
' ****************** 		Zedonius 27/05/2017		  ********************
'                   /***********************************\									|												
'                       **************/\***************
'
'
'
' **************************************************************************
'                                  Credits

'					   Javier15: For all work done
'                      LoboTomy: For the playfield redraw
'                      FBX: For the HR plastic pics 
'                      Akiles: For the HR pics  
'                      Makuste for the all work in VP SCENE
'                      Koadic: New Plunger code 
'                      JPSalas: code, textures, Sounds etc
'                      mfuegemann: models 3d nell and saw
'                   I apologize if I forget to mention someone.
'
' *************************************************************************


option explicit

Randomize

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Const cGameName = "rab_320"
Const BallSize = 54
Const UseSolenoids  = True
Const UseLamps      = False
Const UseSync       = False
Const HandleMech    = False
Const SSolenoidOn   = "fx_Solenoidon"
Const SSolenoidOff  = ""
Const SFlipperOn    = "FlipperUp"
Const SFlipperOff   = "FlipperDown"
Const sCoin         = "Coin3"
Const sDtRDrop      = 15

If RockyBullwinkle.ShowDT = False Then
	Ramp62.visible=0
	Ramp63.visible=0
End If

Dim bsTrough, bsVUK, bsLEjet, bsUpperEject, Lnell, mNell, plungerIM
Dim dtRDrop, bump1, bump2, bump3, x

LoadVPM "01560000","DE.VBS",3.38

Sub InitVPM()
     With Controller
         .GameName = cGameName
         .SplashInfoLine = "Rocky & Bullwinkle And Friends" & vbNewLine & "VPX  Javier15 and Zedonius ver. 0.95"
         .HandleMechanics = 0
         .HandleKeyboard = 0
         .ShowDMDOnly = 1
         .ShowFrame = 0
         .ShowTitle = 0
         .Hidden = 0
         If Err Then MsgBox Err.Description
     End With
     On Error Goto 0
     Controller.SolMask(0) = 0
     vpmTimer.AddTimer 4000, "Controller.SolMask(0)=&Hffffffff'" 
     Controller.Run
End Sub

'**** Solenoid ****
solcallback(1)  ="SolTrough"
Solcallback(2)  ="bsTrough.solout"
solcallback(4)  ="bsVuk.SolOut"
Solcallback(5)  ="bsLEjet.solout" 
SolCallback(8)  ="vpmSolSound SoundFX(""Knocker"",DOFKnocker)," 
SolCallBack(9)  ="SolNell"
solcallback(11) ="SolGi"
solcallback(12) ="SolRdiv"
solcallback(13) ="SolPdiv"
solCallback(14) ="vpmSolgate Gate,SoundFX(""Diverter"",DOFContactors),"
solCallback(15) ="SolDrop" '5 Bank
SolCallBack(16) ="SolAutoPlungerIM"
SolCallBack(17)	="vpmSolSound SoundFX(""jet"",DOFContactors),"
SolCallBack(18)	="vpmSolSound SoundFX(""jet"",DOFContactors),"
SolCallBack(19)	="vpmSolSound SoundFX(""jet"",DOFContactors),"
SolCallBack(20) ="vpmSolSound SoundFX(""slingshot"",DOFContactors),"
SolCallBack(21) ="vpmSolSound SoundFX(""slingshot"",DOFContactors),"	
SolCallBack(22) ="AutoPlunge"

'Flashers
Solcallback(25) = "SetLamp 81," '1
Solcallback(26) = "SetFlash 82," '2
Solcallback(27) = "SetLamp 83," '3
Solcallback(28) = "SetLamp 84," '4
SolCallback(29) = "SetLamp 85," '5	
Solcallback(30) = "SetLamp 86," '6
SolCallback(31) = "SetLamp 87," '7	
SolCallback(32) = "SetLamp 88," '8


 Sub RockyBullwinkle_Init
	 InitVPM
     vpmNudge.TiltSwitch = 1
	 vpmNudge.Sensitivity = 5 
	 vpmNudge.tiltobj = Array(LeftSlingShot,RightSlingShot,Bumper1b,Bumper2b,Bumper3b)


    PinMAMETimer.Interval = PinMAMEInterval
    PinMAMETimer.Enabled = 1

	'Drain & BallRelease
	Set bsTrough = New cvpmBallStack
    With bsTrough
	    .InitSw 0,13,12,11,0,0,0,0
	    .InitKick BallRelease, 90, 10
        .InitExitSnd SoundFX("fx_kickerout",DOFContactors), SoundFX("fx_solenoid",DOFContactors)
	    .Balls = 3
    End With

    ' Left vuk
    Set bsVuk = New cvpmBallStack
    With bsVuk
        .InitSaucer sw52, 52, 218, 58
        .KickZ = 1.5
        .InitExitSnd SoundFX("fx_kickerout",DOFContactors), "kicker_enter"
    End With
 

     'SuperVuk
     Set bsLEjet =new cvpmBallStack
     With  bsLEjet
           .InitSw 0,29,0,0,0,0,0,0
           .InitKick Sw29,177,30
           .KickZ = 0.4
	       .InitExitSnd SoundFX("fx_kickerout",DOFContactors), "Solenoid"
      End With 

     ' Drop Targets
     Set dtRDrop = new cvpmDropTarget
     With dtRDrop
	      .Initdrop Array(Sw17,Sw18,Sw19,Sw20,Sw21), Array(17,18,19,20,21)
	      .InitSnd SoundFX("fx_droptarget",DOFDropTargets),SoundFX("fx_resetdrop",DOFContactors)
      End With 


    ' Impulse Plunger
    Const IMPowerSetting = 54 'Plunger Power
    Const IMTime = 0.6        ' Time in seconds for Full Plunge
    Set plungerIM = New cvpmImpulseP
    With plungerIM
        .InitImpulseP swplunger, IMPowerSetting, IMTime
        .Random 0.3
        .switch 14
        .InitExitSnd "fx_plunger", "fx_plunger"
        .CreateEvents "plungerIM"
    End With
  
     'Diverters & Sensors
     Diverter_on.isdropped=true
	 Diverter_off.isdropped=False
     PlayFLDiverter.isdropped=True

     'Gi
' 	For each x in GILampFlashT:x.isvisible = 0:next
' 	For each x in GILampFlashM:x.isvisible = 0:next
' 	For each x in GILampFlashB:x.isvisible = 0:next 

End sub

'Trough
Sub SolTrough(Enabled)
    If Enabled then
		bsTrough.ExitSol_On
		vpmTimer.PulseSw 10
	End If
End Sub

'AutoPlunger
Sub SolAutoPlungerIM(Enabled)
    If Enabled Then
        PlungerIM.AutoFire
    End If
End Sub


' LaserKick
Sub AutoPlunge(Enabled)
	If Enabled Then 
		LaserKick.Enabled=True
	Else
		LaserKick.Enabled=False
	End If
End Sub

Sub LaserKick_Hit:Me.Kick 0,52:PlaySound SoundFX("fx_bumper2",DOFContactors) End Sub


'Nell Animation
Dim NellPos, NellDir
NellPos = 130
NellDir = 0

Sub SolNell(Enabled)
	
    If Enabled Then
	   NelLight1.state = 1
	   NellDir = -2
       NellTimer.Enabled=1 
      Else
	   NelLight1.state = 0
	   NellDir = 2
       NellTimer.Enabled=1  
  End If
End Sub

Sub NellTimer_Timer()
    StopSound"motor": PlaySound SoundFX("motor",DOFGear) 
	NellPos = NellPos + NellDir
	If NellPos <= 0 Then NellPos = 0:Me.Enabled = 0
	If NellPos >= 130 Then NellPos = 130: Me.Enabled = 0
	NellP.TransY = NellPos
End Sub

' **********
' Diverters
' **********

Dim Diverter2Pos, Diverter2Dir

Diverter2Dir = 0
Diverter2Pos = 0

 Sub SolRdiv(enabled)
 	if enabled then
		Diverter2Animation.Interval = 4
		Diverter2Dir = -1
		Diverter2Animation.Enabled = 1
 		Diverter_on.isdropped=false
 		Diverter_off.isdropped=true
        PlaySound SoundFX("diverter",DOFContactors)
 	else
		Diverter2Animation.Interval = 4
		Diverter2Dir = 1
		Diverter2Animation.Enabled = 1
  		Diverter_on.isdropped=true
 		Diverter_off.isdropped=false
        PlaySound SoundFX("diverter",DOFContactors)
 	end if
 end sub

Sub Diverter2Animation_Timer
    Diverter2.RotY = Diverter2Pos
    Diverter2Pos = Diverter2Pos + Diverter2Dir
    If Diverter2Pos > 0 Then
        Diverter2Pos = 0
    End If
    If Diverter2Pos < -32 Then
        Diverter2Pos = -32
    End If
End Sub

Dim Diverter1Pos, Diverter1Dir

Diverter1Dir = 0
Diverter1Pos = 0

 Sub SolPdiv(enabled)
 	if enabled then
		Diverter1Animation.Interval = 4
		Diverter1Dir = 1
		Diverter1Animation.Enabled = 1
 		PlayFLDiverter.isdropped=false
 		PlayFLDiverter1.isdropped=True
        PlaySound SoundFX("diverter",DOFContactors)
 	else
		Diverter1Animation.Interval = 4
		Diverter1Dir = -1
		Diverter1Animation.Enabled = 1
 		PlayFLDiverter.isdropped=true
 		PlayFLDiverter1.isdropped=false
        PlaySound SoundFX("diverter",DOFContactors)
 	end if
 end sub

Sub Diverter1Animation_Timer
    Diverter1.RotY = Diverter1Pos
    Diverter1Pos = Diverter1Pos + Diverter1Dir
    If Diverter1Pos < 0 Then
        Diverter1Pos = 0
    End If
    If Diverter1Pos > 24 Then
        Diverter1Pos = 24
    End If
End Sub

' **** Key_Up ***
Sub RockyBullwinkle_KeyUp(ByVal keyCode)

	If keycode = LeftFlipperKey Then  Controller.Switch(15) = False  End If
 	If keycode = RightFlipperKey Then	Controller.Switch(16) = False End If	
	If KeyUpHandler(keyCode) Then Exit Sub
	If keyCode=PlungerKey Then  Controller.Switch(9)=false: END IF
End Sub

' **** Key_Down ****
Sub RockyBullwinkle_KeyDown(ByVal keyCode)
   
If keycode = LeftFlipperKey Then  Controller.Switch(15) = True End If
	If keycode = RightFlipperKey Then	Controller.Switch(16) = True End If
    If KeyDownHandler(keyCode) Then Exit Sub
	If keyCode=PlungerKey Then Controller.Switch(9)=true END IF
End Sub



'********************
'     Flippers
'********************

 
 SolCallback(sLRFlipper) = "SolLRFlipper"
 SolCallback(sLLFlipper) = "SolLLFlipper"

 Sub SolLLFlipper(Enabled)
     If Enabled Then
         PlaySound SoundFX("fx_flipup",DOFFlippers),0,1,-0.05,0.25:LeftFlipper.RotateToEnd
     Else
         PlaySound SoundFX("fx_flipdown",DOFFlippers),0,1,-0.05,0.25:LeftFlipper.RotateToStart
 	End If
 End Sub
 
 Sub SolLRFlipper(Enabled)
     If Enabled Then
         PlaySound SoundFX("fx_flipup",DOFFlippers),0,1,0.05,0.25:RightFlipper.RotateToEnd
     Else
         PlaySound SoundFX("fx_flipdown",DOFFlippers),0,1,0.05,0.25:RightFlipper.RotateToStart
     End If
 End Sub
 
Sub LeftFlipper_Collide(parm)
    PlaySound "rubber_hit_3", 0, parm / 10, -0.05, 0.25
End Sub

Sub RightFlipper_Collide(parm)
    PlaySound "rubber_hit_3", 0, parm / 10, 0.05, 0.25
End Sub
  
Sub UpdateModelFlipper_Timer
	LeftBat.RotY=LeftFlipper.currentangle-90
	RightBat.RotY=RightFlipper.currentangle-90
End Sub

'******************
' RealTime Updates
'******************
Set MotorCallback = GetRef("GameTimer")

Sub GameTimer
    RollingSound
	UpdateModelFlipper
End Sub


'**** Drain ****
 Sub Drain_Hit:Playsound "fx_drain":bsTrough.AddBall Me:End Sub
 Sub Drain1_Hit:Playsound "fx_drain":bsTrough.AddBall Me:End Sub
 Sub Drain2_Hit:Playsound "fx_drain":bsTrough.AddBall Me:End Sub
 Sub Drain3_Hit:Playsound "fx_drain":bsTrough.AddBall Me:End Sub
 Sub Drain4_Hit:Playsound "fx_drain":bsTrough.AddBall Me:End Sub

'**** Shooter Lane **** 
Sub Sw14_Hit():Playsound "sensor":Sw14b.WidthBottom = 0:Sw14b.WidthTop = 0:lSw14b.State = 1:Controller.Switch(14)=1: End Sub
Sub Sw14_UnHit():Sw14b.WidthBottom = 4:Sw14b.WidthTop = 4:lSw14b.State = 0:Controller.Switch(14)=0: End Sub

' **** Small Traget ****
Sub sw25_Hit:vpmTimer.PulseSw 25:PlaySound SoundFX("bump",DOFTargets):sw25p.transY = -10:Me.TimerEnabled = 1:End Sub
Sub sw25_Timer:sw25p.transY = 0:Me.TimerEnabled = 0:End Sub

Sub sw26_Hit:vpmTimer.PulseSw 26:PlaySound SoundFX("bump",DOFTargets):sw26p.transY = -10:Me.TimerEnabled = 1:End Sub
Sub sw26_Timer:sw26p.transY = 0:Me.TimerEnabled = 0:End Sub

Sub sw27_Hit:vpmTimer.PulseSw 27:PlaySound SoundFX("bump",DOFTargets):sw27p.transY = -10:Me.TimerEnabled = 1:End Sub
Sub sw27_Timer:sw27p.transY = 0:Me.TimerEnabled = 0:End Sub
 
Sub sw28_Hit:vpmTimer.PulseSw 28:PlaySound SoundFX("bump",DOFTargets):sw28p.transY = -10:Me.TimerEnabled = 1:End Sub
Sub sw28_Timer:sw28p.transY = 0:Me.TimerEnabled = 0:End Sub   

'***************
'  Slingshots
'***************

Dim LStep, RStep

Sub LeftSlingShot_Slingshot
    vpmTimer.PulseSw 30
    PlaySound SoundFX("RightSlingShot",DOFContactors)
    LeftSling2.Visible = 1
    Lemk.RotX = 26
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 1:LeftSLing4.Visible = 0:LeftSLing3.Visible = 1:Lemk.RotX = 14
        Case 2:LeftSLing3.Visible = 0:LeftSLing2.Visible = 1:Lemk.RotX = 2
        Case 3:LeftSLing2.Visible = 0:Lemk.RotX = -10:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

Sub RightSlingShot_Slingshot
    vpmTimer.PulseSw 38
    PlaySound SoundFX("LeftSlingShot",DOFContactors)
    RightSling2.Visible = 1
    Remk.RotX = 26
    RStep = 0
    RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 1:RightSLing4.Visible = 0:RightSLing3.Visible = 1:Remk.RotX = 14
        Case 2:RightSLing3.Visible = 0:RightSLing2.Visible = 1:Remk.RotX = 2
        Case 3:RightSLing2.Visible = 0:Remk.RotX = -10:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub


'*************
' Targets 
'*************
Dim Sw17pos,Sw18pos,Sw19pos,Sw20pos,Sw21pos
Sub SolDrop(enabled)
	if enabled then
       Sw17pos=50:Sw18pos=50:Sw19pos=50:Sw20pos=50:Sw21pos=50
       Sw17Tr.Enabled=1:Sw18Tr.Enabled=1:Sw19Tr.Enabled=1:Sw20Tr.Enabled=1:Sw21Tr.Enabled=1
	   dtRDrop.DropSol_On
	end if
End Sub

Sub Sw17_Hit:dtRDrop.Hit 1:Sw17pos=0:Sw17T.Enabled=1:End Sub
Sub Sw18_Hit:dtRDrop.Hit 2:Sw18pos=0:Sw18T.Enabled=1:End Sub
Sub Sw19_Hit:dtRDrop.Hit 3:Sw19pos=0:Sw19T.Enabled=1:End Sub
Sub Sw20_Hit:dtRDrop.Hit 4:Sw20pos=0:Sw20T.Enabled=1:End Sub
Sub Sw21_Hit:dtRDrop.Hit 5:Sw21pos=0:Sw21T.Enabled=1:End Sub

'SW17
  Sub Sw17T_Timer()	
  Select Case Sw17pos
        Case 0: Sw17p.z=25:Sw17.IsDropped=1        
        Case 1: Sw17p.z=20
        Case 2: Sw17p.z=15
        Case 3: Sw17p.z=10
        Case 4: Sw17p.z=5
        Case 5: Sw17p.z=0
        Case 6: Sw17p.z=-5
        Case 7: Sw17p.z=-10
        Case 8: Sw17p.z=-15
        Case 9: Sw17p.z=-20
        Case 10: Sw17p.z=-20:Sw17T.Enabled = 0  
End Select
 	 If Sw17pos<10 then Sw17pos=Sw17pos+1
  End Sub

  Sub Sw17Tr_Timer()
  Select Case Sw17pos
        Case 0: Sw17p.z=25:Sw17.IsDropped=0:Sw17Tr.Enabled = 0       
        Case 1: Sw17p.z=20
        Case 2: Sw17p.z=15
        Case 3: Sw17p.z=10
        Case 4: Sw17p.z=5
        Case 5: Sw17p.z=0
        Case 6: Sw17p.z=-5
        Case 7: Sw17p.z=-10
        Case 8: Sw17p.z=-15
        Case 9: Sw17p.z=-20
        Case 10: Sw17p.z=-20   
   End Select
 	If Sw17pos>0 Then Sw17pos=Sw17pos-1
  End Sub

'Sw18
  Sub Sw18T_Timer()	
  Select Case Sw18pos
        Case 0: Sw18p.z=25:Sw18.IsDropped=1        
        Case 1: Sw18p.z=20
        Case 2: Sw18p.z=15
        Case 3: Sw18p.z=10
        Case 4: Sw18p.z=5
        Case 5: Sw18p.z=0
        Case 6: Sw18p.z=-5
        Case 7: Sw18p.z=-10
        Case 8: Sw18p.z=-15
        Case 9: Sw18p.z=-20
        Case 10: Sw18p.z=-20:Sw18T.Enabled = 0  
End Select
 	 If Sw18pos<10 then Sw18pos=Sw18pos+1
  End Sub

  Sub Sw18Tr_Timer()
  Select Case Sw18pos
        Case 0: Sw18p.z=25:Sw18.IsDropped=0:Sw18Tr.Enabled = 0       
        Case 1: Sw18p.z=20
        Case 2: Sw18p.z=15
        Case 3: Sw18p.z=10
        Case 4: Sw18p.z=5
        Case 5: Sw18p.z=0
        Case 6: Sw18p.z=-5
        Case 7: Sw18p.z=-10
        Case 8: Sw18p.z=-15
        Case 9: Sw18p.z=-20
        Case 10: Sw18p.z=-20   
   End Select
 	If Sw18pos>0 Then Sw18pos=Sw18pos-1
  End Sub

'Sw19
  Sub Sw19T_Timer()	
  Select Case Sw19pos
        Case 0: Sw19p.z=19:Sw19.IsDropped=1        
        Case 1: Sw19p.z=20
        Case 2: Sw19p.z=15
        Case 3: Sw19p.z=10
        Case 4: Sw19p.z=5
        Case 5: Sw19p.z=0
        Case 6: Sw19p.z=-5
        Case 7: Sw19p.z=-10
        Case 8: Sw19p.z=-15
        Case 9: Sw19p.z=-20
        Case 10: Sw19p.z=-20:Sw19T.Enabled = 0  
End Select
 	 If Sw19pos<10 then Sw19pos=Sw19pos+1
  End Sub

  Sub Sw19Tr_Timer()
  Select Case Sw19pos
        Case 0: Sw19p.z=25:Sw19.IsDropped=0:Sw19Tr.Enabled = 0       
        Case 1: Sw19p.z=20
        Case 2: Sw19p.z=15
        Case 3: Sw19p.z=10
        Case 4: Sw19p.z=5
        Case 5: Sw19p.z=0
        Case 6: Sw19p.z=-5
        Case 7: Sw19p.z=-10
        Case 8: Sw19p.z=-15
        Case 9: Sw19p.z=-20
        Case 10: Sw19p.z=-20   
   End Select
 	If Sw19pos>0 Then Sw19pos=Sw19pos-1
  End Sub

'Sw20
  Sub Sw20T_Timer()	
  Select Case Sw20pos
        Case 0: Sw20p.z=25:Sw20.IsDropped=1        
        Case 1: Sw20p.z=20
        Case 2: Sw20p.z=15
        Case 3: Sw20p.z=10
        Case 4: Sw20p.z=5
        Case 5: Sw20p.z=0
        Case 6: Sw20p.z=-5
        Case 7: Sw20p.z=-10
        Case 8: Sw20p.z=-15
        Case 9: Sw20p.z=-20
        Case 10: Sw20p.z=-20:Sw20T.Enabled = 0  
End Select
 	 If Sw20pos<10 then Sw20pos=Sw20pos+1
  End Sub

  Sub Sw20Tr_Timer()
  Select Case Sw20pos
        Case 0: Sw20p.z=25:Sw20.IsDropped=0:Sw20Tr.Enabled = 0       
        Case 1: Sw20p.z=20
        Case 2: Sw20p.z=15
        Case 3: Sw20p.z=10
        Case 4: Sw20p.z=5
        Case 5: Sw20p.z=0
        Case 6: Sw20p.z=-5
        Case 7: Sw20p.z=-10
        Case 8: Sw20p.z=-15
        Case 9: Sw20p.z=-20
        Case 10: Sw20p.z=-20   
   End Select
 	If Sw20pos>0 Then Sw20pos=Sw20pos-1
  End Sub

'Sw21
  Sub Sw21T_Timer()	
  Select Case Sw21pos
        Case 0: Sw21p.z=25:Sw21.IsDropped=1        
        Case 1: Sw21p.z=20
        Case 2: Sw21p.z=15
        Case 3: Sw21p.z=10
        Case 4: Sw21p.z=5
        Case 5: Sw21p.z=0
        Case 6: Sw21p.z=-5
        Case 7: Sw21p.z=-10
        Case 8: Sw21p.z=-15
        Case 9: Sw21p.z=-20
        Case 10: Sw21p.z=-20:Sw21T.Enabled = 0  
End Select
 	 If Sw21pos<10 then Sw21pos=Sw21pos+1
  End Sub

  Sub Sw21Tr_Timer()
  Select Case Sw21pos
        Case 0: Sw21p.z=25:Sw21.IsDropped=0:Sw21Tr.Enabled = 0       
        Case 1: Sw21p.z=20
        Case 2: Sw21p.z=15
        Case 3: Sw21p.z=10
        Case 4: Sw21p.z=5
        Case 5: Sw21p.z=0
        Case 6: Sw21p.z=-5
        Case 7: Sw21p.z=-10
        Case 8: Sw21p.z=-15
        Case 9: Sw21p.z=-20
        Case 10: Sw21p.z=-20   
   End Select
 	If Sw21pos>0 Then Sw21pos=Sw21pos-1
  End Sub


'***************
' Static Target
'***************

Sub Sw22_Hit: Sw22.IsDropped = 1:Sw22b.IsDropped = 0:Sw22.TimerEnabled = 1:vpmTimer.PulseSwitch(22),0,"" end sub
Sub Sw22_Timer:Sw22.IsDropped = 0:Sw22b.IsDropped = 1:Me.TimerEnabled = 0:End Sub 
Sub Sw23_Hit: Sw23.IsDropped = 1:Sw23b.IsDropped = 0:Sw23.TimerEnabled = 1:vpmTimer.PulseSwitch(23),0,"" end sub
Sub Sw23_Timer:Sw23.IsDropped = 0:Sw23b.IsDropped = 1:Me.TimerEnabled = 0:End Sub   
Sub Sw24_Hit: Sw24.IsDropped = 1:Sw24b.IsDropped = 0:Sw24.TimerEnabled = 1:vpmTimer.PulseSwitch(24),0,"" end sub
Sub Sw24_Timer:Sw24.IsDropped = 0:Sw24b.IsDropped = 1:Me.TimerEnabled = 0:End Sub

'*****************
' Triggers Switch
'*****************


Sub Sw31_Hit():Playsound "sensor":Controller.Switch(31)=1: End Sub
Sub Sw31_UnHit():Controller.Switch(31)=0: End Sub

Sub Sw32_Hit():Playsound "sensor":Controller.Switch(32)=1: End Sub
Sub Sw32_UnHit():Controller.Switch(32)=0: End Sub

Sub Sw33_Hit():PlaySound "fx_gate", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0 :Controller.Switch(33)=1: End Sub
Sub Sw33_UnHit():Controller.Switch(33)=0: End Sub

Sub Sw34_Hit():Playsound "sensor":Controller.Switch(34)=1: End Sub
Sub Sw34_UnHit():Controller.Switch(34)=0: End Sub

Sub Sw35_Hit():Playsound "sensor":Controller.Switch(35)=1: End Sub
Sub Sw35_UnHit():Controller.Switch(35)=0: End Sub

Sub Sw36_Hit():Playsound "sensor":Controller.Switch(36)=1: End Sub
Sub Sw36_UnHit():Controller.Switch(36)=0: End Sub

Sub Sw39_Hit():Playsound "sensor":Controller.Switch(39)=1: End Sub
Sub Sw39_UnHit():Controller.Switch(39)=0: End Sub

Sub Sw40_Hit():Playsound "sensor":Controller.Switch(40)=1: End Sub
Sub Sw40_UnHit():Controller.Switch(40)=0: End Sub

Sub Sw41_Hit():Playsound "sensor":Controller.Switch(41)=1: End Sub
Sub Sw41_UnHit():Controller.Switch(41)=0: End Sub

Sub Sw42_Hit():Playsound "sensor":Controller.Switch(42)=1: End Sub
Sub Sw42_UnHit():Controller.Switch(42)=0: End Sub

Sub Sw43_Hit():Playsound "sensor":Controller.Switch(43)=1: End Sub
Sub Sw43_UnHit():Controller.Switch(43)=0: End Sub

Sub Sw47_Hit(): StopSound "fx_plasticrolling2":PlaySound "fx_gate", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:Controller.Switch(47)=1: End Sub
Sub Sw47_UnHit():Controller.Switch(47)=0: End Sub

Sub Sw48_Hit():ActiveBall.VelY = ActiveBall.VelY / 1.1:Controller.Switch(48)=1: End Sub
Sub Sw48_UnHit():Controller.Switch(48)=0: End Sub

Sub Sw49_Hit(): StopSound "fx_plasticrolling2":PlaySound "fx_gate", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:Controller.Switch(49)=1: End Sub
Sub Sw49_UnHit():Controller.Switch(49)=0: End Sub

Sub Sw50_Hit():Controller.Switch(50)=1: End Sub
Sub Sw50_UnHit():Controller.Switch(50)=0: End Sub

'**** Kickers ****
Sub SW37_Hit():PlaySound "Drain5":SW37.DestroyBall:vpmTimer.PulseSwitch(37),110,"SuperVuk":End Sub
Sub SW37a_Hit():PlaySound "Drain5":SW37a.DestroyBall:vpmTimer.PulseSwitch(37),110,"SuperVuk":End Sub
Sub SW37a1_Hit():PlaySound "Drain5":SW37a1.DestroyBall:vpmTimer.PulseSwitch(37),110,"SuperVuk":End Sub
Sub SW37a2_Hit():PlaySound "Drain5":SW37a2.DestroyBall:vpmTimer.PulseSwitch(37),110,"SuperVuk":End Sub
Sub SW37a3_Hit():PlaySound "Drain5":SW37a3.DestroyBall:vpmTimer.PulseSwitch(37),110,"SuperVuk":End Sub
Sub SW37a4_Hit():PlaySound "Drain5":SW37a4.DestroyBall:vpmTimer.PulseSwitch(37),110,"SuperVuk":End Sub
Sub SuperVuk(SwNo):bsLEjet.AddBall 0:End Sub

Sub Sw29_Hit:Playsound "Kicker_Enter": bsLEjet.AddBall Me:End Sub
'Sub Sw29a_Hit:Playsound "Kicker_Enter" :bsLEjet.AddBall Me:End Sub

Sub sw52_Hit:Playsound "kicker_enter":bsVuk.AddBall 0:End Sub




'**********
' Bumpers
'**********


      Sub Bumper1b_Hit:vpmTimer.PulseSw 44:PlaySound SoundFX("fx_leftbumper",DOFContactors):bump1 = 1:Me.TimerEnabled = 1:End Sub
     
      Sub Bumper2b_Hit:vpmTimer.PulseSw 45:PlaySound SoundFX("fx_topbumper",DOFContactors):bump2 = 1:Me.TimerEnabled = 1:End Sub
     
      Sub Bumper3b_Hit:vpmTimer.PulseSw 46:PlaySound SoundFX("fx_rightbumper",DOFContactors):bump3 = 1:Me.TimerEnabled = 1:End Sub


'*************************************
'          Nudge System
' based on Noah's nudgetest table
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
    If NudgeEffect = 0 then NudgeTimer.Enabled = False
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
LampTimer.Interval = 40 'lamp fading speed
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
End Sub

Sub NFadeLFOm(nr)
    Select Case FadingLevel(nr)
        Case 4:SetFlash nr, 0
        Case 5:SetFlash nr, 1
    End Select
End Sub

Sub UpdateLamps
   'General Lamps
   'SpinLights
   NFadeLm 1, Lamp1
   NFadeLm 2, Lamp2
   NFadeLm 3, Lamp3
   NFadeLm 4, Lamp4
   NFadeLm 5, Lamp5
   NFadeLm 6, Lamp6
   NFadeLm 7, Lamp7
   NFadeLm 8, Lamp8
   NFadeLm 9, Lamp9
   NFadeLm 10, Lamp10
   NFadeLm 11, Lamp11
   NFadeLm 11, Lamp12
   NFadeLm 11, Lamp13
   NFadeLm 11, Lamp14

   'Dropp Targets
   NFadeLm 17, Lamp17
   NFadeLm 18, Lamp18
   NFadeLm 19, Lamp19
   NFadeLm 20, Lamp20
   NFadeLm 21, Lamp21

   NFadeLm 22, Lamp22
   NFadeLm 23, Lamp23
   NFadeLm 24, Lamp24
   NFadeLm 25, Lamp25
   NFadeLm 26, Lamp26
   NFadeLm 27, Lamp27
   NFadeLm 28, Lamp28

   'LaserKick Millions
   NFadeLm 29, Lamp29
   NFadeLm 30, Lamp30
   NFadeLm 31, Lamp31

   
   NFadeLm 32, Lamp32
   NFadeLm 33, Lamp33
   NFadeLm 34, Lamp34
   NFadeLm 35, Lamp35
   NFadeLm 36, Lamp36
   NFadeLm 37, Lamp37
   NFadeLm 38, Lamp38
   NFadeLm 39, Lamp39
   NFadeLm 40, Lamp40

   'Daisy Select
   NFadeLm 41, Lamp41
   NFadeLm 42, Lamp42
   NFadeLm 43, Lamp43
   NFadeLm 44, Lamp44
   NFadeLm 45, Lamp45
   NFadeLm 46, Lamp46

   NFadeLm 47, Lamp47
  ' FadeL 48, Lamp48, Lamp48a
   NFadeLm 49, Lamp49
   NFadeLm 50, Lamp50
   NFadeLm 51, Lamp51
   NFadeLm 52, Lamp52
'   FadeL 53, Lamp53, Lamp53a
   NFadeLm 54, Lamp54
   NFadeLm 55, Lamp55
   NFadeLm 56, Lamp56

   'Hat Target
   NFadeLm 57, Lamp57
   NFadeLm 58, Lamp58
   NFadeLm 59, Lamp59

   NFadeLm 60, Lamp60
   NFadeLm 61, Lamp61
   NFadeLm 62, Lamp62
   NFadeLm 63, Lamp63
   NFadeLm 64, Lamp64
   
   ' Flashers 
   NFadeLm 16, l16
   NFadeLm 53, l53
   NFadeLm 85, l5r
   NFadeLm 81,  Gi1r1
   NFadeLFOm 82
   NFadeLm 82,  GiR2
   NFadeLFOm 83
   FadeL 83, GiR3a, GiR3b
   NFadeLm 84, Gi4r1
'   NFadeLm 84, Gi4r2
   NFadeLm 86, Gi6r1
   NFadeL2 87, Gi7r1
   NFadeLm 88, Gi8r1
'   NFadePri 85, bulb8, "Bulb_On", "Bulb_Off"

	'GI LIGHTS
   NFadeLm 101, GiL1
   NFadeLm 102, GiL2
   NFadeLm 103, GiL3
   NFadeLm 104, GiL4
   NFadeLm 105, GiL5
   NFadeLm 106, GiL6
   NFadeLm 107, GiL7
   NFadeLm 108, GiL8
   NFadeLm 109, GiL9
   NFadeLm 110, GiL10
   NFadeLm 111, GiL11
   NFadeLm 112, GiL12
   NFadeLm 113, GiL13
   NFadeLm 114, GiL14
   NFadeLm 115, GiL15
   NFadeLm 116, GiL16
   NFadeLm 117, GiL17
   NFadeLm 118, GiL18
   NFadeLm 119, GiL19
   NFadeLm 120, GiL20
   NFadeLm 121, GiL21
   NFadeLm 122, GiL22
   NFadeLm 123, GiL23
   NFadeLm 124, GiL24
   NFadeLm 125, GiL25
   NFadeLm 126, GiL26
   NFadeLm 127, GiL27
   NFadeLm 128, GiL28
   NFadeLm 129, GiL29
   NFadeLm 130, GiL30
   NFadeLm 131, GiL31
   NFadeLm 132, GiL32
   NFadeLm 133, GiL33
   NFadeLm 134, GiL34
   NFadeLm 135, GiL35
   NFadeLm 136, GiL36
   NFadeLm 137, GiL37
   NFadeLm 138, GiL38
   NFadeLm 139, GiL39
   NFadeLm 140, GiL40
   NFadeLm 141, GiL41
   NFadeLm 142, GiL42

If Lamp1.state = 1 Then
		l1.visible = 1
		l1b.visible = 1
		ToyL1.visible = 1
		ToyL1b.visible = 0
	Else
		l1.visible = 0
		l1b.visible = 0
		ToyL1.visible = 0
		ToyL1b.visible = 1
end if
If Lamp2.state = 1 Then
		l2.visible = 1
		l2b.visible = 1
		ToyL2.visible = 1
		ToyL2b.visible = 0
	Else
		l2.visible = 0
		l2b.visible = 0
		ToyL2.visible = 0
		ToyL2b.visible = 1
end if
If Lamp3.state = 1 Then
		l3.visible = 1
		l3b.visible = 1
		ToyL3.visible = 1
		ToyL3b.visible = 0
	Else
		l3.visible = 0
		l3b.visible = 0
		ToyL3.visible = 0
		ToyL3b.visible = 1
end if
If Lamp4.state = 1 Then
		l4.visible = 1
		l4b.visible = 1
		ToyL4.visible = 1
		ToyL4b.visible = 0
	Else
		l4.visible = 0
		l4b.visible = 0
		ToyL4.visible = 0
		ToyL4b.visible = 1
end if
If Lamp5.state = 1 Then
		l5.visible = 1
		l5b.visible = 1
		ToyL5.visible = 1
		ToyL5b.visible = 0
	Else
		l5.visible = 0
		l5b.visible = 0
		ToyL5.visible = 0
		ToyL5b.visible = 1
end if
If Lamp6.state = 1 Then
		l6.visible = 1
		l6b.visible = 1
		ToyL6.visible = 1
		ToyL6b.visible = 0
	Else
		l6.visible = 0
		l6b.visible = 0
		ToyL6.visible = 0
		ToyL6b.visible = 1
end if
If Lamp7.state = 1 Then
		l7.visible = 1
		l7b.visible = 1
		ToyL7.visible = 1
		ToyL7b.visible = 0
	Else
		l7.visible = 0
		l7b.visible = 0
		ToyL7.visible = 0
		ToyL7b.visible = 1
end if
If Lamp8.state = 1 Then
		l8.visible = 1
		l8b.visible = 1
		ToyL8.visible = 1
		ToyL8b.visible = 0
	Else
		l8.visible = 0
		l8b.visible = 0
		ToyL8.visible = 0
		ToyL8b.visible = 1
end if
If Lamp9.state = 1 Then
		l9.visible = 1
		l9b.visible = 1
		ToyL9.visible = 1
		ToyL9b.visible = 0
	Else
		l9.visible = 0
		l9b.visible = 0
		ToyL9.visible = 0
		ToyL9b.visible = 1
end if
If Lamp10.state = 1 Then
		l10.visible = 1
		l10b.visible = 1
		ToyL10.visible = 1
		ToyL10b.visible = 0
	Else
		l10.visible = 0
		l10b.visible = 0
		ToyL10.visible = 0
		ToyL10b.visible = 1
end if
If Lamp11.state = 1 Then
		l11.visible = 1
		l11b.visible = 1
		ToyL11.visible = 1
		ToyL11b.visible = 0
	Else
		l11.visible = 0
		l11b.visible = 0
		ToyL11.visible = 0
		ToyL11b.visible = 1
end if
If Lamp12.state = 1 Then
		l12.visible = 1
		l12b.visible = 1
		ToyL12.visible = 1
		ToyL12b.visible = 0
	Else
		l12.visible = 0
		l12b.visible = 0
		ToyL12.visible = 0
		ToyL12b.visible = 1
end if
If Lamp13.state = 1 Then
		l13.visible = 1
		l13b.visible = 1
		ToyL13.visible = 1
		ToyL13b.visible = 0
	Else
		l13.visible = 0
		l13b.visible = 0
		ToyL13.visible = 0
		ToyL13b.visible = 1
end if
If Lamp14.state = 1 Then
		l14.visible = 1
		l14b.visible = 1
		ToyL14.visible = 1
		ToyL14b.visible = 0
	Else
		l14.visible = 0
		l14b.visible = 0
		ToyL14.visible = 0
		ToyL14b.visible = 1
end if

If Gi7r1.state = 1 Then
		Gi7r2.state = 1
		Gi7r3.state = 1
		Gi7r2b.visible = 1
	Else
		Gi7r2.state = 0
		Gi7r3.state = 0
		Gi7r2b.visible = 0
end if

If l5r.state = 1 Then
		l5rb.visible = 1
		l5rb1.visible = 1
		l5rb2.visible = 1
		l5rc.visible = 1
		l5rc1.visible = 1
	Else
		l5rb.visible = 0
		l5rb1.visible = 0
		l5rb2.visible = 0
		l5rc.visible = 0
		l5rc1.visible = 0
end if

If l16.state = 1 Then
		l16b.visible = 1
		l16b1.visible = 1
		l16b2.visible = 1
		LightLR1.visible = 0
		LightLR2.visible = 1
	Else
		l16b.visible = 0
		l16b1.visible = 0
		l16b2.visible = 0
		LightLR1.visible = 1
		LightLR2.visible = 0
end if
If l53.state = 1 Then
		l53b.visible = 1
		l53b1.visible = 1
		l53b2.visible = 1
		LightRR1.visible = 0
		LightRR2.visible = 1
	Else
		l53b.visible = 0
		l53b1.visible = 0
		l53b2.visible = 0
		LightRR1.visible = 1
		LightRR2.visible = 0
end if

   'Gi Lights
End Sub


' div lamp subs

Sub AllLampsOff()
    Dim x
    For x = 0 to 200
        LampState(x) = 0
        FadingLevel(x) = 4
    Next

UpdateLamps:UpdateLamps:Updatelamps
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

    FlashSpeedUp = 50   ' fast speed when turning on the flasher
    FlashSpeedDown = 10 ' slow speed when turning off the flasher, gives a smooth fading
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

' div lamp & flasher subs for each kind of VP object

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

Sub FadeL(nr, a, b)
    Select Case FadingLevel(nr)
        Case 2:b.state = 0:FadingLevel(nr) = 0
        Case 3:b.state = 1:FadingLevel(nr) = 2
        Case 4:a.state = 0:FadingLevel(nr) = 3
        Case 5:a.state = 1:FadingLevel(nr) = 1
    End Select
End Sub

Sub FadeLm(nr, a, b)
    Select Case FadingLevel(nr)
        Case 2:b.state = 0
        Case 3:b.state = 1
        Case 4:a.state = 0
        Case 5:a.state = 1
    End Select
End Sub

Sub NFadeL(nr, a)
    Select Case FadingLevel(nr)
        Case 4:a.state = 0:FadingLevel(nr) = 0
        Case 5:a.State = 1:FadingLevel(nr) = 1
    End Select
End Sub

Sub NFadeL2(nr, a)
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

' Flasher objects
' Uses own faster timer

'***********
' GILights
'***********

Sub SolGi(enabled)
Dim GIState
If enabled then
	GIState = 0
Bumper1L.visible = 0
Bumper2L.visible = 0
Bumper3L.visible = 1
Bumper3L1.visible = 0
Bumper3L2.visible = 0
    SetLamp 101, 0
    SetLamp 102, 0
    SetLamp 103, 0
    SetLamp 104, 0
    SetLamp 105, 0
    SetLamp 106, 0
    SetLamp 107, 0
    SetLamp 108, 0
    SetLamp 109, 0
    SetLamp 110, 0
    SetLamp 111, 0
    SetLamp 112, 0
    SetLamp 113, 0
    SetLamp 114, 0
    SetLamp 115, 0
    SetLamp 116, 0
    SetLamp 117, 0
    SetLamp 118, 0
    SetLamp 119, 0
    SetLamp 120, 0
    SetLamp 121, 0
    SetLamp 122, 0
    SetLamp 123, 0
    SetLamp 124, 0
    SetLamp 125, 0
    SetLamp 126, 0
    SetLamp 127, 0
    SetLamp 128, 0
    SetLamp 129, 0
    SetLamp 130, 0
    SetLamp 131, 0
    SetLamp 132, 0
    SetLamp 133, 0
    SetLamp 134, 0
    SetLamp 135, 0
    SetLamp 136, 0
    SetLamp 137, 0
    SetLamp 138, 0
    SetLamp 139, 0
    SetLamp 140, 0
    SetLamp 141, 0
    SetLamp 142, 0
else
	GIState = 1
Bumper1L.visible = 1
Bumper2L.visible = 1
Bumper3L.visible = 0
Bumper3L1.visible = 1
Bumper3L2.visible = 1
    SetLamp 101, 1
    SetLamp 102, 1
    SetLamp 103, 1
    SetLamp 104, 1
    SetLamp 105, 1
    SetLamp 106, 1
    SetLamp 107, 1
    SetLamp 108, 1
    SetLamp 109, 1
    SetLamp 110, 1
    SetLamp 111, 1
    SetLamp 112, 1
    SetLamp 113, 1
    SetLamp 114, 1
    SetLamp 115, 1
    SetLamp 116, 1
    SetLamp 117, 1
    SetLamp 118, 1
    SetLamp 119, 1
    SetLamp 120, 1
    SetLamp 121, 1
    SetLamp 122, 1
    SetLamp 123, 1
    SetLamp 124, 1
    SetLamp 125, 1
    SetLamp 126, 1
    SetLamp 127, 1
    SetLamp 128, 1
    SetLamp 129, 1
    SetLamp 130, 1
    SetLamp 131, 1
    SetLamp 132, 1
    SetLamp 133, 1
    SetLamp 134, 1
    SetLamp 135, 1
    SetLamp 136, 1
    SetLamp 137, 1
    SetLamp 138, 1
    SetLamp 139, 1
    SetLamp 140, 1
    SetLamp 141, 1
    SetLamp 142, 1
end If
	SetFlash 101, GIState
	SetFlash 102, GIState
	SetFlash 103, GIState
	SetFlash 104, GIState
	SetFlash 105, GIState
	SetFlash 106, GIState
	SetFlash 107, GIState
	SetFlash 108, GIState
	SetFlash 109, GIState
	SetFlash 110, GIState
	SetFlash 111, GIState
	SetFlash 112, GIState
	SetFlash 113, GIState
	SetFlash 114, GIState
	SetFlash 115, GIState
	SetFlash 116, GIState
	SetFlash 117, GIState
	SetFlash 118, GIState
	SetFlash 119, GIState
	SetFlash 120, GIState
	SetFlash 121, GIState
	SetFlash 122, GIState
	SetFlash 123, GIState
	SetFlash 124, GIState
	SetFlash 125, GIState
	SetFlash 126, GIState
	SetFlash 127, GIState
	SetFlash 128, GIState
	SetFlash 129, GIState
	SetFlash 130, GIState
	SetFlash 131, GIState
	SetFlash 132, GIState
	SetFlash 133, GIState
	SetFlash 134, GIState
	SetFlash 135, GIState
	SetFlash 136, GIState
	SetFlash 137, GIState
	SetFlash 138, GIState
	SetFlash 139, GIState
	SetFlash 140, GIState
	SetFlash 141, GIState
	SetFlash 142, GIState

End Sub

'** End Extra math **

'Algunos Efectos

Sub Balldrop1_Hit: PlaySound "fx_ballrampdrop": End Sub
Sub Balldrop2_Hit: PlaySound "fx_ballrampdrop": End Sub
Sub Balldrop3_Hit: PlaySound "rubber": End Sub

Sub LRFXOn_Hit
	PlaySound "fx_metalrolling", 0, 0.2
End Sub

Sub LRFXOff_Hit
	StopSound "fx_metalrolling"
End Sub

Sub RRFXON_Hit
	PlaySound "fx_metalrolling", 0, 0.2
End Sub

Sub RRFXOff_Hit
	StopSound "fx_metalrolling"
End Sub

Sub RRPLasticFXON_Hit
	PlaySound "fx_plasticrolling2", 0, 0.2, Pitch(ActiveBall)
End Sub

Sub RRPlasticFXOff_Hit
	StopSound "fx_plasticrolling2"
End Sub

Sub RRPLasticFXON1_Hit
	PlaySound "fx_plasticrolling1", 0, 0.2, Pitch(ActiveBall)
End Sub

Sub RRPlasticFXOff1_Hit
	StopSound "fx_plasticrolling1"
End Sub

Sub RRPLasticFXON2_Hit
	PlaySound "fx_plasticrolling2", 0, 0.2, Pitch(ActiveBall)
End Sub

Sub RRPlasticFXOff2_Hit
	StopSound "fx_plasticrolling2"
End Sub

Sub RRPLasticFXON3_Hit
	PlaySound "fx_plasticrolling1", 0, 0.2, Pitch(ActiveBall)
End Sub



Sub aRubbers_Hit(idx):PlaySound "fx_rubber", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
Sub aPostRubbers_Hit(idx):PlaySound "fx_postrubber", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub

' *********************************************************************
'                      Supporting Ball & Sound Functions
' *********************************************************************

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
    Vol = Csng(BallVel(ball) ^2 / 500)
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "table1" is the name of the table
    Dim tmp
    tmp = ball.x * 2 / RockyBullwinkle.width-1
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

Const tnob = 7 ' total number of balls in this table is 4, but always use a higher number here bacuse of the timing
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
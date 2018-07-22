 '*********************************************************************
 '*                                                                   *
 '*                           Doctor  who                             *
 '*                            Script by                              *
 '*                    oooPLAYER1ooo & Unclewilly                     *
 '*                               2010                                *
 '*                                    								  *
 '*                         Updated 2017 for VPX By                   *
 '*								Sliderpoint                           *
 '*********************************************************************

' Thalamus 2018-07-20
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Changed UseSolenoids=1 to 2
' No special SSF tweaks yet.
' This table doesn't contain the standard subs I normally add.
 
   Option Explicit
   Randomize
 
On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Sub LoadCoreVBS
 On Error Resume Next
 ExecuteGlobal GetTextFile("core.vbs")
 If Err Then MsgBox "Can't open core.vbs"
 On Error Goto 0
End Sub
  
Const UseVPMModSol = 1

LoadVPM "01560000", "WPC.VBS", 3.26
 
   '********************
   'Standard definitions
   '********************
   
	Const UseSolenoids = 2
	Const UseLamps = 1
	Const UseSync = 0
	Set GICallback2 = GetRef("UpdateGI")

   ' Standard Sounds
   Const SSolenoidOn = "Solenoid"
   Const SSolenoidOff = ""
   Const SCoin = "quarter"
   
  'Rom Name
    Const cGameName = "dw_l2"


'XXXXXXXXXXXXX - Graphics Variables - XXXXXXXXXXX  
   Const GI_Color = "White" ' Mixed - Red - Blue - White
   Const SideWallFlashers = 1 ' 1 On / 0 Off
   Const GISideWalls = 1 ' 1 On / 0 Off
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX 

 'Table Init
   Dim bsTrough,bsMiniPFL,bsMiniPFR,bsShooter,mMinipf,plungerIM
   Dim PFPos
   Dim xx
   Dim DayNight

   Sub Table1_Init
       vpmInit Me
       With Controller
  		  .GameName = cGameName
           If Err Then MsgBox "Can't start Game " & cGameName & vbNewLine & Err.Description:Exit Sub
           .SplashInfoLine = "Doctor Who" & vbNewLine & "by Sliderpoint v0.5 vp10.2"
           .HandleKeyboard = 0
           .ShowTitle = 0
           .ShowDMDOnly = 1
           .ShowFrame = 0
           .HandleMechanics = 0
           .Hidden = 0
  		   .dip(0)=&h00  'Set to usa
		   Controller.Switch(22) = 1 'close coin door
		   Controller.Switch(24) = 1 'always closed
		   Controller.Switch(82) = 1 'pfglass switch
           On Error Resume Next
           .Run 
 
           If Err Then MsgBox Err.Description
           On Error Goto 0
        End With
  
' Impulse Plunger
    Const IMPowerSetting = 45
    Const IMTime = 0.7
    Set plungerIM = New cvpmImpulseP
    With plungerIM
        .InitImpulseP swplunger, IMPowerSetting, IMTime
        .InitExitSnd SoundFX("Solenoid",DOFContactors), SoundFX("plunger",DOFContactors)
        .CreateEvents "plungerIM"
    End With

 ' Ballstacks
  	set bstrough=new cvpmballstack
  	bstrough.initsw 28,25,26,27,0,0,0,0
  	bstrough.initkick ballrelease,0,1
 	bstrough.InitExitSnd SoundFX("BallRelease",DOFContactors), SoundFX("solenoid",DOFContactors)
	bstrough.balls=3

  	set mMinipf=new cvpmmech
  	with mMinipf
  		.sol1=28
  		.sol2=27
  		.mtype=vpmMechOneDirSol+vpmmechlinear
  		.length=270
  		.steps=360
  		.callback=getref("UpdateMiniPF")
  		.start
  	end with
  
 'Other Init
	sw71.isDropped = 1
	sw72.isDropped = 1
	sw73.isDropped = 1
	sw74.isDropped = 1
	sw75.isDropped = 1

	DayNight = table1.NightDay
	Intensity 'sets GI brightness depending on day/night slider settings

 'Init switches
   Controller.Switch(22) = 1 'close coin door
   Controller.Switch(24) = 1 'always closed
   Controller.Switch(82) = 1 'pfglass switch
 
	vpmNudge.TiltSwitch = 14
	vpmNudge.Sensitivity = 4
	vpmNudge.TiltObj = Array(Bumper1, Bumper2, Bumper4, LeftSling, RightSling)  

 'Init Trapdoor
 	tdenter.enabled=true
		
'Graphic Variables
If GI_Color = "Mixed" then
	for each xx in GIG2:xx.Color = RGB(255, 0, 0):next
	for each xx in GIG3:xx.Color = RGB(255, 255, 0):next
	for each xx in GIG4:xx.Color = RGB(0, 0, 255):next
	for each xx in GIG2:xx.ColorFull = RGB(255, 0, 0):next
	for each xx in GIG3:xx.ColorFull = RGB(255, 255, 0):next
	for each xx in GIG4:xx.ColorFull = RGB(0, 0, 255):next
	for each xx in GIF2:xx.Color = RGB(255, 255, 128):Next
	for each xx in GIF3:xx.Color = RGB(255, 255, 0):Next
	for each xx in GIF4:xx.Color = RGB(0, 0, 255):Next

End If

If GI_Color = "Red" then
	for each xx in GIG2:xx.Color = RGB(255, 0, 0):next
	for each xx in GIG3:xx.Color = RGB(255, 0, 0):next
	for each xx in GIG4:xx.Color = RGB(255, 0, 0):next
	for each xx in GIG2:xx.ColorFull = RGB(255, 0, 0):next
	for each xx in GIG3:xx.ColorFull = RGB(255, 0, 0):next
	for each xx in GIG4:xx.ColorFull = RGB(255, 0, 0):next
	for each xx in GIF2:xx.Color = RGB(255, 0, 0):next
	for each xx in GIF3:xx.Color = RGB(255, 0, 0):next
	for each xx in GIF4:xx.Color = RGB(255, 0, 0):next
End If

If GI_Color = "White" then
	for each xx in GIG2:xx.Color = RGB(255, 255, 0):next
	for each xx in GIG3:xx.Color = RGB(255, 255, 0):next
	for each xx in GIG4:xx.Color = RGB(255, 255, 0):next
	for each xx in GIF2:xx.Color = RGB(255, 255, 128):Next
	for each xx in GIF3:xx.Color = RGB(255, 255, 128):Next
	for each xx in GIF4:xx.Color = RGB(255, 255, 128):Next
End If

If GI_Color = "Blue" then
	for each xx in GIG2:xx.Color = RGB(0, 0, 255):next
	for each xx in GIG3:xx.Color = RGB(0, 0, 255):next
	for each xx in GIG4:xx.Color = RGB(0, 0, 255):next
	for each xx in GIG2:xx.ColorFull = RGB(0, 0, 255):next
	for each xx in GIG3:xx.ColorFull = RGB(0, 0, 255):next
	for each xx in GIG4:xx.ColorFull = RGB(0, 0, 255):next
	for each xx in GIF2:xx.Color = RGB(0, 0, 255):Next
	for each xx in GIF3:xx.Color = RGB(0, 0, 255):Next
	for each xx in GIF4:xx.Color = RGB(0, 0, 255):Next
End If

End Sub
	
PFPos=-1

Sub Table1_Paused:Controller.Pause = 1:End Sub
Sub Table1_unPaused:Controller.Pause = 0:End Sub 
Sub Table1_exit:Controller.Stop: End Sub
 
  '*********************keyboard handlers**************************
 
   Sub Table1_KeyDown(ByVal Keycode)
  	 If keycode = plungerkey then controller.switch(34)=True
 '    If keycode = 3 then SolTrapDoor(1)
    If keycode = LeftTiltKey Then PlaySound "fx_nudge_left"
    If keycode = RightTiltKey Then PlaySound "fx_nudge_right"
    If keycode = CenterTiltKey Then :PlaySound "fx_nudge_forward"     
    If vpmKeyDown(keycode) Then Exit Sub 
    End Sub
   
   Sub Table1_KeyUp(ByVal Keycode)
        If vpmKeyUp(keycode) Then Exit Sub
        If keyuphandler(keycode) Then Exit Sub
  	 If keycode = plungerkey then controller.switch(34)=false
'     If keycode = 3 then SolTrapDoor(0)
   End Sub
 
 'Ball events
sub Drain_hit():PlaySoundAt "Drain", Drain:bsTrough.AddBall me:end sub
sub sw77_hit():controller.switch(77) = 1:PlaySoundAt "scoopenter",sw77:BallPrim2.visible = 1:end sub
Sub sw77_unhit:controller.Switch(77) = 0:end sub
sub sw76_hit():controller.switch(76) = 1:PlaySoundAt "scoopenter",sw76:BallPrim.visible = 1:end sub 
Sub sw76_unhit:controller.switch(76) = 0:end Sub
sub TardisEntrance_hit:Controller.Switch(31) = 1:end sub
Sub ShooterLane_Hit:Controller.Switch(17)=1:End Sub
Sub ShooterLane_Unhit:Controller.Switch(17)=0:End Sub

  'MiniPF Door Switches
sub sw68_Hit:vpmTimer.PulseSw 68:End Sub
Sub Pins_Hit (idx)
	PlaySoundAtBall "pinhit_low"
End Sub

Sub Targets_Hit (idx)
	PlaySoundAtBall "target"
End Sub

Sub Metals_Thin_Hit (idx)
	RePlaySoundAtBall "metalhit_thin"
End Sub

Sub Metals_Medium_Hit (idx)
	RePlaySoundAtBall "metalhit2"
End Sub

Sub Metals2_Hit (idx)
	REPlaySoundAtBall "metalhit"
End Sub

Sub Gates_Hit (idx)
	RePlaySoundAtBall "gate4"
End Sub

Sub Rubbers_Hit(idx)
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 20 then 
		PlaySoundAtBall "fx_rubber2"
	End if
	If finalspeed >= 6 AND finalspeed <= 20 then
 		RandomSoundRubber()
 	End If
End Sub

Sub RubberPosts_Hit(idx)
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 16 then 
		PlaySoundAtBall "fx_rubber2"
	End if
	If finalspeed >= 6 AND finalspeed <= 16 then
 		RandomSoundRubber()
 	End If
End Sub

Sub RandomSoundRubber()
	Select Case Int(Rnd*3)+1
		Case 1 : PlaySoundAtBall "rubber_hit_1"
		Case 2 : PlaySoundAtBall "rubber_hit_2"
		Case 3 : PlaySoundAtBall "rubber_hit_3"
	End Select
End Sub


sub sw38_Hit:vpmTimer.PulseSw 38:End Sub
sub sw38s_Hit:PlaysoundAt "scoopenter",sw38s: End Sub
sub sw88_Hit:vpmTimer.PulseSw 88:End Sub
sub sw38s_Hit:PlaysoundAt "scoopenter",sw38s: End Sub


 'MiniPF Standup
Sub sw78_Hit():vpmTimer.PulseSw 78:PlaySound "target":End Sub
 
 'MiniPf Buttons
Sub sw75_Hit():vpmTimer.PulseSw 75:PlaySound "target":sw75p.Y = 440:TESHake:ButtonPrim.Enabled = 1:End Sub
Sub sw74_Hit():vpmTimer.PulseSw 74:PlaySound "target":sw74p.Y = 440:TESHake:ButtonPrim.Enabled = 1:End Sub
Sub sw73_Hit():vpmTimer.PulseSw 73:PlaySound "target":sw73p.Y = 440:TESHake:ButtonPrim.Enabled = 1:End Sub
Sub sw72_Hit():vpmTimer.PulseSw 72:PlaySound "target":sw72p.Y = 440:TESHake:ButtonPrim.Enabled = 1:End Sub
Sub sw71_Hit():vpmTimer.PulseSw 71:PlaySound "target":sw71p.Y = 440:TESHake:ButtonPrim.Enabled = 1:End Sub
 
Sub ButtonPrim_Timer
	sw75p.Y = 455
	sw74p.Y = 455
	sw73p.Y = 455
	sw72p.Y = 455
	sw71p.Y = 455
 Me.Enabled = 0
End Sub

 'ramp gates
Sub sw36_Hit:vpmTimer.PulseSw 36:PlaySoundAt "gate",sw36:End Sub
Sub gate3_Hit:vpmTimer.PulseSw 37:PlaySoundAt "gate",gate3:End Sub
Sub sw33_Hit:vpmTimer.PulseSw 33:PlaySoundAt "gate",sw33:End Sub
Sub gate5_Hit:vpmTimer.PulseSw 35:PlaySoundAt "gate",gate5:End Sub
Sub gate2_Hit:PlaySoundAt "gate",gate2:End Sub
  
 ' Activate transmat
Sub sw58_Hit:vpmTimer.PulseSw 58:PlaySoundAt "target",sw58:End Sub
 
 ' Escape targets
Sub sw41_Hit:vpmTimer.PulseSw 41:PlaySoundAt "target",sw41:End Sub
Sub sw42_Hit:vpmTimer.PulseSw 42:PlaySoundAt "target",sw42:End Sub
Sub sw43_Hit:vpmTimer.PulseSw 43:PlaySoundAt "target",sw43:End Sub
Sub sw44_Hit:vpmTimer.PulseSw 44:PlaySoundAt "target",sw44:End Sub
Sub sw45_Hit:vpmTimer.PulseSw 45:PlaySoundAt "target",sw45:End Sub
Sub sw46_Hit:vpmTimer.PulseSw 46:PlaySoundAt "target",sw46:End Sub
 
 'repair targets
Sub sw51_Hit:vpmTimer.PulseSw 51:PlaySoundAt "target",sw51:End Sub
Sub sw52_Hit:vpmTimer.PulseSw 52:PlaySoundAt "target",sw52:End Sub
Sub sw53_Hit:vpmTimer.PulseSw 53:PlaySoundAt "target",sw53:End Sub
Sub sw54_Hit:vpmTimer.PulseSw 54:PlaySoundAt "target",sw54:End Sub
Sub sw55_Hit:vpmTimer.PulseSw 55:PlaySoundAt "target",sw55:End Sub
Sub sw56_Hit:vpmTimer.PulseSw 56:PlaySoundAt "target",sw56:End Sub
 
 ' lane rollovers
Sub RightOutlane_Hit:Controller.Switch(67) = 1:PlaySoundAt "sensor",RightOutlane:End Sub
Sub RightOutlane_UnHit:Controller.Switch(67) = 0:End Sub
Sub RightInlane_Hit:Controller.Switch(66) = 1:PlaySoundAt "sensor",RightInlane:End Sub
Sub RightInlane_UnHit:Controller.Switch(66) = 0:End Sub
Sub LeftOutlane_Hit:Controller.Switch(64) = 1:PlaySoundAt "sensor",LeftOutlane:End Sub
Sub LeftOutlane_UnHit:Controller.Switch(64) = 0:End Sub
Sub LeftInlane_Hit:Controller.Switch(65) = 1:PlaySoundAt "sensor",LeftInlane:End Sub
Sub LeftInlane_UnHit:Controller.Switch(65) = 0:End Sub
Sub LeftMiddle_Hit:Controller.Switch(47) = 1:PlaySoundAt "sensor",LeftMiddle:End Sub
Sub LeftMiddle_UnHit:Controller.Switch(47) = 0:End Sub
   
 'hidden rollovers
Sub sw18_Hit:vpmTimer.PulseSw 18:PlaySoundAt "sensor",sw18:End Sub
Sub sw48_Hit:vpmTimer.PulseSw 48:PlaySoundAt "sensor",sw48:End Sub
 
' slings
Sub leftsling_Slingshot():vpmTimer.PulseSw 15:PlaySound SoundFX("slingshot_L" ,DOFContactors):End Sub
Sub rightsling_Slingshot():vpmTimer.PulseSw 16:PlaySound SoundFX("slingshot_R" ,DOFContactors):End Sub

 'Bumpers
Dim Bump1,Bump2,Bump3

Sub Bumper2_Hit():vpmTimer.PulseSw 61:PlaySound SoundFX("" ,DOFContactors):PlaySoundAt "Bumper1",bumper2:End Sub
Sub Bumper1_Hit():vpmTimer.PulseSw 62:PlaySound SoundFX("" ,DOFContactors):PlaySoundAt "Bumper2",bumper1:End Sub
Sub Bumper4_Hit():vpmTimer.PulseSw 63:PlaySound SoundFX("" ,DOFContactors):PlaySoundAt "Bumper3",bumper4:End Sub'
 
 'Solenoids
solcallback(1)="SolTrapDoor"
solcallback(2)="SolAutoFire"
solcallback(3)="TardisExit"
solcallback(4)="solmpfl" 
solcallback(5)="solmpfr"
solModcallback(6)= "Flash06"
solcallback(7)="vpmsolsound ""knocker"","
'solcallback(8)= 'doctor 3 flasher, in backbox
'solcallback(11)="bpr1"
'solcallback(12)="bpr2"
'solcallback(13)="bpr3"
solcallback(15)="bsTrough.SolIn"
solcallback(16)="bsTrough.SolOut"
solModcallback(17)="Flash17"
solModcallback(18)="Flash18"
solModcallback(19)="Flash19"
solModcallback(20)="flash20"
solModcallback(21)="Flash21"
solModCallback(22)= "who_h"
solModCallback(23)= "who_o"
solModcallback(24)="Flash24"
  
'Solenoid Subs
sub soltrapdoor(Enabled)
 	if enabled then
 		tdenter.enabled= false
		TDDownTimer.Enabled = 1
		Playsound SoundFX("FlapClos", DOFContactors), 0,.25,-1
 	else
 		tdenter.enabled= true
		TDUpTimer.Enabled = 1
	end if
  end sub

Sub TDUpTimer_timer
		If TD.RotX > 20 Then
			TD.RotX = TD.RotX - 1
		elseIf TD.RotX = 20 Then
 		controller.switch(57)= False
			TDUpTimer.Enabled = 0
		Playsound SoundFX("FlapClos", DOFContactors), 0,.25,-1,1
		End If
End Sub

Sub TDDownTimer_timer
		If TD.RotX < 55 Then
			TD.RotX = TD.RotX + 1
		elseIf TD.RotX = 55 Then
 		controller.switch(57)= True
			TDDownTimer.Enabled = 0
		End If
End Sub

Sub solAutofire(Enabled)
	If Enabled Then
		PlungerIM.AutoFire
        Playsound SoundFX("Solenoid" ,DOFContactors)
	End If
End Sub
  
Sub TardisExit(enabled)
	If Enabled Then
		TardisEntrance.KickZ 180, 35, 92, 0
        Playsound SoundFX("FlapOpen" ,DOFContactors),0,1,.5,.5
		Controller.Switch (31) = 0
	End If
End Sub

Sub solmpfl(enabled)
 	If enabled then
		BallPrim.visible = 0
		sw76.kick 180, 5
		Playsound SoundFX("BallRelease2" ,DOFContactors)
 	End If
End Sub
 
Sub solmpfr(enabled)
	If enabled then
		BallPrim2.visible = 0
		sw77.kick 180, 5
		Playsound SoundFX("BallRelease2" ,DOFContactors)
	End If
End Sub

Sub Flash06(Level)
	If Level > 0 Then
		FL06.IntensityScale = (Level / 2.55) / 100
		FL06b.IntensityScale = (Level / 2.55) / 100
		FL06.State = 1
		FL06b.State = 1
	Else
		FL06.State = 0
		FL06b.State = 0
	End If
End Sub

Sub Flash17(Level)
	If Level > 0 Then
		FL17.IntensityScale = (Level / 2.55) / 100
		FL17.State = 1
		TEFlashP.image = "TopFlasherRedOn"
	Else
		FL17.State = 0
		TEFlashP.image = "TopFlasherRed"
	End If
End Sub

Sub Flash18(Level)
	If Level > 0 Then
		FL18.IntensityScale = (Level / 2.55) / 100
		FL18.State = 1
	Else
		FL18.State = 0
	End If
End Sub

Sub Flash19(Level)
	If Level > 0 Then
		FL19.IntensityScale = (Level / 2.55) / 100
		FL19.State = 1
	Else
		FL19.State = 0
	End If
End Sub

Sub Flash20(Level)
	If Level > 0 Then
		FL20.Opacity = (Level / 2.55)
		FL20.visible = 1
	 else
		FL20.Visible = 0
	 end if
end sub

Sub Flash21(Level)
	If Level > 0 Then	
		FL21.IntensityScale = (Level / 2.55) / 100
		FL21b.IntensityScale = (Level / 2.55) / 100
		FL21.State = 1
		FL21b.State = 1
		If SideWallFlashers = 1 then
			FL21c.visible = 1
			FL21d.visible = 1
		End If
	Else
		FL21.State = 0
		FL21b.State = 0
		FL21c.visible = 0
		FL21d.visible = 0

	End If
End Sub

Sub who_h(Level)
	If Level > 0 Then
		FL22.IntensityScale = (Level / 2.55) / 100
		FL22.State = 1
     else
		FL22.State = 0
	end if
End Sub
 
Sub who_o(Level)
	If Level > 0 Then
		FL23.IntensityScale = (Level / 2.55) / 100
		FL23.State = 1
     else
		FL23.State = 0
     end if
End Sub

Sub Flash24(Level)
	If Level > 0 Then
		FL24.IntensityScale = (Level / 2.55) / 100
		FL24.State = 1
	Else
		FL24.State = 0
	End If
End Sub


  '**************
  ' Flipper Subs
  '**************
SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"
  
  Sub SolLFlipper(Enabled)
      If Enabled Then
          PlaySound SoundFX("" ,DOFFlippers),0,1,-.5:LeftFlipper.RotateToEnd:LeftFlipper2.RotateToEnd
		  PlaySoundAt "flipperup", LeftFlipper
      Else
          PlaySound SoundFX("" ,DOFFlippers),0,1,-.5:LeftFlipper.RotateToStart:LeftFlipper2.RotateToStart
		  PlaySoundAt "FlipperDown", LeftFlipper
      End If
  End Sub
  
  Sub SolRFlipper(Enabled)
      If Enabled Then
          PlaySound SoundFX("" ,DOFFLippers),0,1,.5:RightFlipper.RotateToEnd
		  PlaySoundAt "flipperup", RightFlipper
      Else
          PlaySound SoundFX("" ,DOFFlippers),0,1,.5:RightFlipper.RotateToStart
		  PlaySoundAt "FlipperDown", RightFlipper
      End If
  End Sub
 
 
'**********************************************************
'     MiniPF Animation 
 '**********************************************************
Dim ZPos

Sub UpdateMiniPF(aCurrPos,aSpeed,aLast)

		If aCurrPos > 180 Then
			ZPos = (((aCurrPos - 180)* -1) +180)
		Else 
			ZPos = aCurrPos
		End If
			
		For Each XX in MiniPF
			XX.TransZ = (ZPos * .7843)
			Playsound SoundFX("Motor-Old1" ,DOFGear),0,0.1,0,.1
		Next

		For Each XX in MiniPF2
			XX.Z = ((ZPos * .7843) - 84.18)
		Next
		For Each XX in MiniPFLights
			XX.BulbHaloHeight = (ZPos * .7843)
		Next

  Dim OldLevel :OldLevel = PFPos
  	If aCurrPos < 35 then Controller.Switch(32) = false
 	If aCurrPos > 35 and aCurrPos < 90 then Controller.Switch(32) = true
 	If aCurrPos > 90 and aCurrPos < 145 then Controller.Switch(32) = false
 	If aCurrPos > 145 and aCurrPos < 180 then Controller.Switch(32) = true
 	If aCurrPos > 180 and aCurrPos < 190 then Controller.Switch(32) = false
 	If aCurrPos > 190 and aCurrPos < 270 then Controller.Switch(32) = true
 	If aCurrPos > 270 and aCurrPos < 350 then Controller.Switch(32) = false
 	If aCurrPos > 350then Controller.Switch(32) = true
		If aCurrPos < 46 then PFPos = 0
		If aCurrPos > 45 and aCurrPos < 136 then PFPos = 1
		If aCurrPos > 135 and aCurrPos < 226 then PFPos = 2
		If aCurrPos > 225 and aCurrPos < 316 then PFPos = 1
		If aCurrPos > 315 then PFPos = 0
 			If OldLevel <> PFPos Then	
				Select Case PFPos
						Case 0:'Ground level
								Gate68.Collidable = 0
								Gate38.Collidable = 0
								Gate88.Collidable = 0
								sw68s.Enabled = 0
								sw38s.Enabled = 0
								sw88s.Enabled = 0
								sw71.isDropped = 1
								sw72.isDropped = 1
								sw73.isDropped = 1
								sw74.isDropped = 1
								sw75.isDropped = 1
								Pin1.Collidable = 0
								Pin2.Collidable = 0
								Pin3.Collidable = 0.
								Pin4.Collidable = 0
								Wall20.IsDropped = 1
								Wall21.IsDropped = 1
								Scoop3.Collidable = 0
								Scoop4.Collidable = 0
								Scoop5.Collidable = 0
						Case 1:'Buttons
								sw71.isDropped = 0
								sw72.isDropped = 0
								sw73.isDropped = 0
								sw74.isDropped = 0
								sw75.isDropped = 0
						Case 2:'Doors
								Gate68.Collidable = 1
								Gate38.Collidable = 1
								Gate88.Collidable = 1
								sw68s.Enabled = 1
								sw38s.Enabled = 1
								sw88s.Enabled = 1
								sw71.isDropped = 1
								sw72.isDropped = 1
								sw73.isDropped = 1
								sw74.isDropped = 1
								sw75.isDropped = 1
								Pin1.Collidable = 1
								Pin2.Collidable = 1
								Pin3.Collidable = 1
								Pin4.Collidable = 1
								Wall20.IsDropped = 0
								Wall21.IsDropped = 0
								Scoop3.Collidable = 1
								Scoop4.Collidable = 1
								Scoop5.Collidable = 1
				End Select
		End If
End Sub

Sub TEHit_hit (idx)
		TEShake
End Sub

Sub TEShake
		Playsound "metalhit2"
 		For Each XX in MiniPF
			XX.TransY = (Pitch(activeball) * -.01)/2
		Next
		For Each XX in MiniPF2
			XX.TransY = (Pitch(activeball) * -.01)/2
		Next
		ResetMPF.Enabled = 1
End Sub

Sub ResetMPF_timer
 		For Each XX in MiniPF
			XX.TransY = 0
		Next
		For Each XX in MiniPF2
			XX.TransY = 0
		Next
		me.Enabled = 0
End Sub

  '**********************G.I STRING********************************
Sub UpdateGI(giNo, status)
Dim ii
   Select Case giNo
'		Case 0  'BackBox 1,Insert in ROM

'		Case 1  'BackBox 2,Insert in ROM

		Case 2  'String 3,PFa/Inserta in ROM
		If Status > 0 Then
			For each xx in GIG2 
			xx.State = 1
			xx.IntensityScale = Status * .25: Next
			If GISideWalls = 1 then
			For each xx in GIF2
			xx.visible = 1: Next
			end if
		Else
			For each xx in GIG2: xx.State = 0: Next
			For each xx in GIF2: xx.visible = 0: Next
		End If

		Case 3 'String 4,PFb/Insertb in ROM
		If Status > 0 Then
			For each xx in GIG3 
			xx.State = 1
			xx.IntensityScale = Status * .25: Next
			If GISideWalls = 1 then
				For each xx in GIF3
				xx.visible = 1: Next
			End If
		Else
			For each xx in GIG3: xx.State = 0: Next
			For each xx in GIF3: xx.visible = 0:Next
		End If

		Case 4  'String 5,PFc/Insertc in ROM
		If Status > 0 Then
			For each xx in GIG4 
			xx.State = 1
			xx.IntensityScale = Status * .25: Next
			If GISideWalls = 1 then
				For each xx in GIF4
				xx.visible = 1: Next
			End if
		Else
			For each xx in GIG4: xx.State = 0: Next
			For each xx in GIF4: xx.visible = 0: Next
		End If

'		Case 5 	'never passed from ROM 'PFc/Insertc in ROM
	End Select

End Sub

Dim GILevel

Sub Intensity
	If DayNight <= 20 Then 
			GILevel = .5
	ElseIf DayNight <= 40 Then 
			GILevel = .4125
	ElseIf DayNight <= 60 Then 
			GILevel = .325
	ElseIf DayNight <= 80 Then 
			GILevel = .2375
	Elseif DayNight <= 100  Then 
			GILevel = .15
	End If

	For each xx in GIG2: xx.Intensity = xx.Intensity * GILevel: Next
	For each xx in GIG3: xx.Intensity = xx.Intensity * GILevel: Next
	For each xx in GIG4: xx.Intensity = xx.Intensity * GILevel: Next
End Sub

'*************Basic Insert Lights*****************
'****set the timer internval to be the ROM Lamp number on the light Object
'*********add lights to the AllLights collection, Create a LampTimer Timer object.

	vpmMapLights AllLights
'****************************************** 
 
' misc 
Sub PrimTimer_Timer
	gate68p.RotX = Gate68.currentAngle
	Gate88p.RotX = gate88.CurrentAngle
	Gate38p.RotX = Gate38.CurrentAngle
	If Controller.Lamp(67) = 0 Then
		l67on.visible = 0
			If SideWallFlashers = 1 then
			l67onb.visible = 0
			End If
	Else
		L67on.visible = 1
			If SideWallFlashers = 1 then		
			l67onb.visible = 1
			End If
	End If
		If SideWallFlashers = 1 then
			If GI33.State = 1 Then
				GI33b.visible = 1
			Else
				GI33b.visible =0
			End If
			If GI32.State = 1 Then
				GI32b.visible = 1
			Else
				GI32b.visible = 0
			End If
		End if
	If Light3.state = 1 Then
		Flasher4.visible = 1
	Else
		Flasher4.visible = 0
	End If

	If Light2.state = 1 Then
		Flasher5.visible = 1
		Flasher6.visible = 1
	Else
		Flasher5.visible = 0
		Flasher6.visible = 0
	End If
'TextBox2.Text = ABS(gi11.Intensity)
'TextBox1.Text = 
End Sub

' *********************************************************************
'                      Supporting Ball & Sound Functions
' *********************************************************************
Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
    Vol = Csng(BallVel(ball) ^2 / 2000)
End Function

Function Vol2(ball1, ball2) ' Calculates the Volume of the sound based on the speed of two balls
    Vol2 = (Vol(ball1) + Vol(ball2) ) / 2
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "table1" is the name of the table
    Dim tmp
    tmp = ball.x * 2 / table1.width-1
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
'    JP's VP10 Collision & Rolling Sounds
'*****************************************
Const tnob = 8 ' maximum number of balls on the table (including locked/newton Balls)
ReDim rolling(tnob)
ReDim collision(tnob)
Initcollision

Sub Initcollision
    Dim i
    For i = 0 to tnob
        collision(i) = -1
        rolling(i) = False
    Next
End Sub

Sub CollisionTimer_Timer()
    Dim BOT, B, B1, B2, dx, dy, dz, distance, radii
    BOT = GetBalls

' rolling
	For B = UBound(BOT) +1 to tnob
        rolling(b) = False
        StopSound("fx_ballrolling" & b)
	Next

    If UBound(BOT) = -1 Then Exit Sub

    For B = 0 to UBound(BOT)
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

 'collision
    If UBound(BOT) < 1 Then Exit Sub

    For B1 = 0 to UBound(BOT)
        For B2 = B1 + 1 to UBound(BOT)
            dz = INT(ABS((BOT(b1).z - BOT(b2).z) ) )
            radii = BOT(b1).radius + BOT(b2).radius
			If dz <= radii Then

            dx = INT(ABS((BOT(b1).x - BOT(b2).x) ) )
            dy = INT(ABS((BOT(b1).y - BOT(b2).y) ) )
            distance = INT(SQR(dx ^2 + dy ^2) )

            If distance <= radii AND (collision(b1) = -1 OR collision(b2) = -1) Then
                collision(b1) = b2
                collision(b2) = b1
                PlaySound("fx_collide"), 0, Vol2(BOT(b1), BOT(b2)), Pan(BOT(b1)), 0, Pitch(BOT(b1)), 0, 0
            Else
                If distance > (radii + 10)  Then
                    If collision(b1) = b2 Then collision(b1) = -1
                    If collision(b2) = b1 Then collision(b2) = -1
                End If
            End If
			End If
        Next
    Next
End Sub

	'ball drops
Sub RHelp_Hit:PlaySoundAt "DROP_RIGHT",RHelp:StopSound "WireRamp":End Sub' 'ActiveBall.VelY=0
Sub RHelp2_Hit:ActiveBall.VelY=0:PlaySound "DROP_LEFT":End Sub
Sub LaneEnd1_hit:Playsound "rubber_hit_2",0,1,1,.5 :end Sub
Sub LaneEnd2_hit:Playsound "rubber_hit_2",0,1,-1,.5 :end Sub
Sub Kicker4_hit:Playsound "Trough1",0,1,.8,.5 :end Sub
Sub Trigger2_Hit:PlaySound "WireRamp",0,1,.75,.1:end Sub
Sub tdenter_hit:playsound "DROP_LEFT",0,.5,0:end sub


Sub Pins_Hit (idx)
	PlaySoundAtBall "pinhit_low"
End Sub

Sub Targets_Hit (idx)
	PlaySoundAtBall "target"
End Sub

Sub Metals_Thin_Hit (idx)
	RePlaySoundAtBall "metalhit_thin"
End Sub

Sub Metals_Medium_Hit (idx)
	RePlaySoundAtBall "metalhit2"
End Sub

Sub Metals2_Hit (idx)
	RePlaySoundAtBall "metalhit"
End Sub

Sub Gates_Hit (idx)
	RePlaySoundAtBall "gate4"
End Sub

Sub Rubbers_Hit(idx)
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 20 then 
		PlaySoundAtBall "fx_rubber2"
	End if
	If finalspeed >= 6 AND finalspeed <= 20 then
 		RandomSoundRubber()
 	End If
End Sub

Sub RubberPosts_Hit(idx)
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 16 then 
		PlaySoundAtBall "fx_rubber2"
	End if
	If finalspeed >= 6 AND finalspeed <= 16 then
 		RandomSoundRubber()
 	End If
End Sub

Sub RandomSoundRubber()
	Select Case Int(Rnd*3)+1
		Case 1 : PlaySoundAtBall "rubber_hit_1"
		Case 2 : PlaySoundAtBall "rubber_hit_2"
		Case 3 : PlaySoundAtBall "rubber_hit_3"
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
		Case 1 : PlaySoundAtBall "flip_hit_1"
		Case 2 : PlaySoundAtBall "flip_hit_2"
		Case 3 : PlaySoundAtBall "flip_hit_3"
	End Select
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

function AudioFade(ball)
	Dim tmp
    tmp = ball.y * 2 / Table1.height-1
    If tmp > 0 Then
		AudioFade = Csng(tmp ^10)
    Else
        AudioFade = Csng(-((- tmp) ^10) )
    End If
End Function

Function RndNum(min,max)
 RndNum = Int(Rnd()*(max-min+1))+min     ' Sets a random number between min and max
End Function

'Set position as table object (Use object or light but NOT wall) and Vol to 1
Sub PlaySoundAt(sound, tableobj)
	PlaySound sound, 1, 1, Pan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

'Set all as per ball position & speed.
Sub PlaySoundAtBall(sound)
	PlaySound sound, 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 1, AudioFade(ActiveBall)
End Sub

'PlaySound "metalhit2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
Sub RePlaySoundAtBall(sound)
	PlaySound sound, 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
End Sub

'Set position as table object and Vol manually.
Sub PlaySoundAtVol(sound, tableobj, Vol)
	PlaySound sound, 1, Vol, Pan(tableobj), 0, 0, 0, 1, AudioFade(tableobj)
End Sub

Sub PlayLoopSoundAtVol(sound, tableobj, Vol)
	PlaySound sound, -1, Vol, Pan(tableobj), 0, 0, 1, 0, AudioFade(tableobj)
End Sub

'Set all as per ball position & speed, but Vol Multiplier may be used eg; PlaySoundAtBallVol "sound",3
Sub PlaySoundAtBallVol(sound, VolMult)
	PlaySound sound, 0, Vol(ActiveBall) * VolMult, Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 1, AudioFade(ActiveBall)
End Sub

'Set position as bumperX and Vol manually.
Sub PlaySoundAtBumperVol(sound, tableobj, Vol)
	PlaySound sound, 1, Vol, Pan(tableobj), 0,0,1, 1, AudioFade(tableobj)
End Sub

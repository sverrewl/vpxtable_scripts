'* 
'
' Credits:
'  Solenoids and switch information from  "THE SIX MILLION DOLLAR MAN, Version 1.2"
' LoadVPM code from "SharkeysShootout"
' Primitive Target Code from "Tag Team"
'

Option Explicit
Randomize

Dim xx, Bump1, Bump2, Bump3
Dim bsTrough, bsSaucer
Dim dtR
Dim cGameName
Dim LStep,RStep
Dim droptarget1pos,droptarget2pos,droptarget3pos,droptarget4pos,droptarget5pos
droptarget1pos=0:droptarget2pos=0:droptarget3pos=0:droptarget4pos=0:droptarget5pos=0
  
LoadVPM "01560000", "bally.vbs", 3.26  

'******************* Options *********************
' DMD/Backglass Controller Setting (by gtxjoe)
Const cController = 3		'0=Use value defined in cController.txt, 1=VPinMAME, 2=UVP server, 3=B2S server, 4=B2S with DOF (disable VP mech sounds)
'*************************************************

Dim cNewController
       
Sub LoadVPM(VPMver, VBSfile, VBSver)
	Dim FileObj, ControllerFile, TextStr

	On Error Resume Next
	If ScriptEngineMajorVersion < 5 Then MsgBox "VB Script Engine 5.0 or higher required"
	ExecuteGlobal GetTextFile(VBSfile)
	If Err Then MsgBox "Unable to open " & VBSfile & ". Ensure that it is in the same folder as this table. " & vbNewLine & Err.Description

	cNewController = 1
	If cController = 0 then
		Set FileObj=CreateObject("Scripting.FileSystemObject")
		If Not FileObj.FolderExists(UserDirectory) then 
			Msgbox "Visual Pinball\User directory does not exist. Defaulting to vPinMame"
		ElseIf Not FileObj.FileExists(UserDirectory & "cController.txt") then
			Set ControllerFile=FileObj.CreateTextFile(UserDirectory & "cController.txt",True)
			ControllerFile.WriteLine 1: ControllerFile.Close
		Else
			Set ControllerFile=FileObj.GetFile(UserDirectory & "cController.txt")
			Set TextStr=ControllerFile.OpenAsTextStream(1,0)
			If (TextStr.AtEndOfStream=True) then
				Set ControllerFile=FileObj.CreateTextFile(UserDirectory & "cController.txt",True)
				ControllerFile.WriteLine 1: ControllerFile.Close
			Else
				cNewController=Textstr.ReadLine: TextStr.Close
			End If
		End If
	Else
		cNewController = cController
	End If

	Select Case cNewController
		Case 1
			Set Controller = CreateObject("VPinMAME.Controller")
			If Err Then MsgBox "Can't Load VPinMAME." & vbNewLine & Err.Description
			If VPMver>"" Then If Controller.Version < VPMver Or Err Then MsgBox "VPinMAME ver " & VPMver & " required."
			If VPinMAMEDriverVer < VBSver Or Err Then MsgBox VBSFile & " ver " & VBSver & " or higher required."
		Case 2
			Set Controller = CreateObject("UltraVP.BackglassServ")
		Case 3,4
			Set Controller = CreateObject("B2S.Server")
	End Select
	On Error Goto 0
End Sub

Function SoundFX (sound)
    If cNewController= 4 Then
        SoundFX = ""
    Else
        SoundFX = sound
    End If
End Function

Sub DOF(dofevent, dofstate)	
	If cNewController>2 Then
		If dofstate = 2 Then
			Controller.B2SSetData dofevent, 1:Controller.B2SSetData dofevent, 0
		Else
			Controller.B2SSetData dofevent, dofstate
		End If
	End If
End Sub

Const UseSolenoids = 1
Const UseLamps = true
Const UseGI = 1
Const UseSync = 0
Const HandleMech = 0 
 
'Standard Sounds
Const SSolenoidOn = "solon"
Const SSolenoidOff = "soloff"
Const SCoin = "CoinIn"
' Lights numbers are kept in the Timer Interval of the Light and the ROM controls them
'set Lights(59)=l59 ' Array(l59,l59b) ' Credit Indicator - bug - need this in the script to work
'set Lights(56)=l56 ' 


 Sub Table1_Init
     With Controller
   		      cGameName = "smmanc" 
               .GameName = cGameName
               .SplashInfoLine = "Six Million Dollar Man, The, Bally 1978" & vbNewLine & "by allknowing2012"
               .HandleMechanics = 0
               .HandleKeyboard = 0
               .ShowDMDOnly = 1
               .ShowFrame = 0
               .ShowTitle = 0
               .Hidden = 1
               If Err Then MsgBox Err.Description
           End With
           On Error Goto 0
           Controller.Run

     'Nudging
      vpmNudge.TiltSwitch = 7
      vpmNudge.Sensitivity = 5
      vpmNudge.TiltObj = Array(sw38, sw39, sw40, LeftSlingshot, RightSlingShot)   ' 3 bumpers and slings.
 
      '**Trough
       Set bsTrough = New cvpmBallStack
       With bsTrough
 			   .InitSw 0,8,0,0,0,0,0,0
      		   .InitKick Spawn, 45, 8 
  			   .InitExitSnd SoundFX("launchball"), SoundFX("Solenoid")
      		   .BallImage = "ballDark"
			   .Balls = 1
       End With
 
      '**Saucer
       Set bsSaucer = New cvpmBallStack
       With bsSaucer
      		   .InitSaucer sw32, 32, 180, 5
 			   .KickForceVar = 2
			   .KickAngleVar = 2
  			   .InitExitSnd SoundFX("scoopexit"), SoundFX("solon")
       End With

   '**Main Timer init
    PinMAMETimer.Interval = PinMAMEInterval
    PinMAMETimer.Enabled = 1
    
    PostUp.IsDropped=True
	
    vpmMapLights AllLights

 'GIBulbs
   for each xx in GI:xx.State=1:next

	Set dtR=New cvpmDropTarget
	dtR.InitDrop Array(sw01,sw02,sw03,sw04,sw05),Array(1,2,3,4,5)
	dtR.InitSnd SoundFX("drop"),SoundFX("dropsup")

End Sub

Sub Table1_Paused:Controller.Pause = 1:End Sub
Sub Table1_unPaused:Controller.Pause = 0:End Sub

Sub Table1_KeyDown(ByVal keycode)
	If keycode = PlungerKey Then
		Plunger.PullBack
		PlaySound "plungerpull",0,1,0.25,0.25
	End If
  
	If keycode = LeftTiltKey Then
		Nudge 90, 2
	End If
    
   ' if keycode = 72 then sw29_hit()    ' numpad 7 post up
    'if keycode = 71 then sw22_hit()   ' numpad 8 post down

	If keycode = RightTiltKey Then
		Nudge 270, 2
	End If
    
	If keycode = CenterTiltKey Then
		Nudge 0, 2
	End If   
    If vpmKeyDown(keycode) Then Exit Sub
End Sub

Sub Table1_KeyUp(ByVal keycode)
    If vpmKeyUp(keycode) Then Exit Sub
	If keycode = PlungerKey Then
		Plunger.Fire
		PlaySound "plunger",0,1,0.25,0.25
	End If
End Sub

'Solenoids
SolCallback(17)="vpmSolDiverter Diverter,SoundFX(""diverter""),"
SolCallback(13)="SolPostUp"
SolCallback(1)="SolPostDown"
SolCallback(7)="bsTrough.SolOut"
SolCallback(6)="vpmSolSound SoundFX(""knocker""),"
SolCallback(12)="vpmSolSound SoundFX(""sling"")," 
SolCallback(14)="vpmSolSound SoundFX(""sling"")," 
SolCallback(9)="vpmSolSound SoundFX(""bumper""),"  
SolCallback(10)="vpmSolSound SoundFX(""bumper""),"  
SolCallback(11)="vpmSolSound SoundFX(""bumper""),"  
SolCallBack(8)="bsSaucer.SolOut"
'SolCallback(15)="dtR.SolDropUp"
SolCallBack(15)="SolDropTargetsRight"
SolCallback(19)="vpmNudge.SolGameOn"
 
'Flipper Subs
SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"
   
Sub SolLFlipper(Enabled)
   If Enabled Then
       PlaySound SoundFX("fx_flipperup"):LeftFlipper.RotateToEnd
   Else
       PlaySound SoundFX("fx_flipperdown"):LeftFlipper.RotateToStart
   End If
End Sub
     
Sub SolRFlipper(Enabled)
   If Enabled Then
       PlaySound SoundFX("fx_flipperup"):RightFlipper.RotateToEnd
   Else
       PlaySound SoundFX("fx_flipperdown"):RightFlipper.RotateToStart
   End If
End Sub

Sub SolPostUp(Enabled)
	If Enabled Then
        PostUp.IsDropped=False
        PopUpStopper.TransY = -20	
		PlaySound SoundFX("soloff")
	End If
End Sub

Sub SolPostDown(Enabled)
	If Enabled Then
		PostUp.IsDropped=True
        PopUpStopper.TransY = -46	
        PlaySound SoundFX("solon")
	End If
End Sub

Sub Drain_Hit()
	PlaySound "drain",0,1,0,0.25
    bsTrough.AddBall Me
End Sub

' Saucer Hit
Sub sw32_Hit:bsSaucer.AddBall 0:End Sub

'**********Sling Shot Animations

Sub RightSlingShot_Slingshot
    PlaySound SoundFX("left_slingshot"), 0, 1, 0.05, 0.05
    vpmTimer.PulseSw 36
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.TransZ = -20
    RStep = 0
    RightSlingShot.TimerEnabled = 1
	gi1.State = LightStateOff:Gi2.State = LightStateOff
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.TransZ = -10
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.TransZ = 0:RightSlingShot.TimerEnabled = 0:gi1.State = LightStateOn:Gi2.State = LightStateOn
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
    PlaySound SoundFX("right_slingshot"),0,1,-0.05,0.05
    vpmTimer.PulseSw 37
    LSling.Visible = 0
    LSling1.Visible = 1
    sling2.TransZ = -20
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
	gi3.State = LightStateOff:Gi4.State = LightStateOff
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.TransZ = -10
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.TransZ = 0:LeftSlingShot.TimerEnabled = 0:gi3.State = LightStateOn:Gi4.State = LightStateOn
    End Select
    LStep = LStep + 1
End Sub

'Spinner
Sub sw20_Spin():vpmTimer.pulsesw 20:Playsound "fx_spinner":End Sub
Sub sw21_Spin():vpmTimer.pulsesw 21:Playsound "fx_spinner":End Sub


'Bumpers
Sub sw38_Hit
	PlaySound SoundFX("fx_bumper4"):B3L1.State = 1:B3L2. State = 1:vpmTimer.pulsesw 38
	Me.TimerEnabled = 1
End Sub

Sub sw38_Timer
	B3L1.State = 0:B3L2. State = 0:Me.Timerenabled = 0
End Sub

Sub sw39_Hit
	PlaySound SoundFX("fx_bumper4"):B2L1.State = 1:B2L2. State = 1:vpmTimer.pulsesw 39
	Me.TimerEnabled = 1
End Sub

Sub sw39_Timer
	B2L1.State = 0:B2L2. State = 0:Me.Timerenabled = 0
End Sub	

Sub sw40_Hit
	PlaySound SoundFX("fx_bumper4"):B1L1.State = 1:B1L2. State = 1:vpmTimer.pulsesw 40
	Me.TimerEnabled = 1
End Sub

Sub sw40_Timer
	B1L1.State = 0:B1L2. State = 0:Me.Timerenabled = 0
End Sub


'Drop Targets
Sub sw01_Hit:dtR.Hit 1:me.timerenabled=1:vpmTimer.PulseSw 1:PlaySound SoundFX("target"):End Sub
Sub sw02_Hit:dtR.Hit 2:me.timerenabled=1:vpmTimer.PulseSw 2:PlaySound SoundFX("target"):End Sub
Sub sw03_Hit:dtR.Hit 3:me.timerenabled=1:vpmTimer.PulseSw 3:PlaySound SoundFX("target"):End Sub
Sub sw04_Hit:dtR.Hit 4:me.timerenabled=1:vpmTimer.PulseSw 4:PlaySound SoundFX("target"):End Sub
Sub sw05_Hit:dtR.Hit 5:me.timerenabled=1:vpmTimer.PulseSw 5:PlaySound SoundFX("target"):End Sub

Sub SolDropTargetsRight(enabled)
	if enabled then
		dtR.DropSol_On
		if droptarget1pos=48 then rdt1.enabled=1
		if droptarget2pos=48 then rdt2.enabled=1
		if droptarget3pos=48 then rdt3.enabled=1
		if droptarget4pos=48 then rdt4.enabled=1
        if droptarget5pos=48 then rdt5.enabled=1
	end if
End Sub

Sub sw01_timer()
	droptarget1pos=droptarget1pos+4
	sw01p.rotandtra4=0-droptarget1pos
	if droptarget1pos=48 then me.timerenabled=0
End Sub
Sub rdt1_timer()
	droptarget1pos=droptarget1pos-12
	sw01p.rotandtra4=0-droptarget1pos
	if droptarget1pos=0 then me.enabled=0
End Sub

Sub sw02_timer()
	droptarget2pos=droptarget2pos+4
    sw02p.rotandtra4=0-droptarget2pos
	if droptarget2pos=48 then me.timerenabled=0
End Sub
Sub rdt2_timer()
	droptarget2pos=droptarget2pos-12
	sw02p.rotandtra4=0-droptarget2pos
	if droptarget2pos=0 then me.enabled=0
End Sub

Sub sw03_timer()
	droptarget3pos=droptarget3pos+4
	sw03p.rotandtra4=0-droptarget3pos
	if droptarget3pos=48 then me.timerenabled=0
End Sub
Sub rdt3_timer()
	droptarget3pos=droptarget3pos-12
	sw03p.rotandtra4=0-droptarget3pos
	if droptarget3pos=0 then me.enabled=0
End Sub

Sub sw04_timer()
	droptarget4pos=droptarget4pos+4
	sw04p.rotandtra4=0-droptarget4pos
	if droptarget4pos=48 then me.timerenabled=0
End Sub
Sub rdt4_timer()
	droptarget4pos=droptarget4pos-12
	sw04p.rotandtra4=0-droptarget4pos
	if droptarget4pos=0 then me.enabled=0
End Sub

Sub sw05_timer()
	droptarget5pos=droptarget5pos+4
	sw05p.rotandtra4=0-droptarget5pos
	if droptarget5pos=48 then me.timerenabled=0
End Sub
Sub rdt5_timer()
	droptarget5pos=droptarget5pos-12
	sw05p.rotandtra4=0-droptarget5pos
	if droptarget5pos=0 then me.enabled=0
End Sub


'Other Targets
Sub sw22_Hit:vpmTimer.PulseSw 22:sw22p.Transx=-6:sw22.TimerEnabled=1:PlaySound SoundFX("target"):End Sub
Sub sw23_Hit:vpmTimer.PulseSw 23:sw23p.Transx=-6:sw23.TimerEnabled=1:PlaySound SoundFX("target"):End Sub
Sub sw24_Hit:vpmTimer.PulseSw 24:sw24p.Transx=-6:sw24.TimerEnabled=1:PlaySound SoundFX("target"):End Sub

Sub sw27_Hit:vpmTimer.PulseSw 27:sw27p.Transx=-6:sw27.TimerEnabled=1:PlaySound SoundFX("target"):End Sub
Sub sw28_Hit:vpmTimer.PulseSw 28:sw28p.Transx=-6:sw28.TimerEnabled=1:PlaySound SoundFX("target"):debug.print "Up/Down":End Sub
Sub sw29_Hit:vpmTimer.PulseSw 29:sw29p.Transx=-6:sw29.TimerEnabled=1:PlaySound SoundFX("target"):debug.print "Up/Down":End Sub

Sub sw22_Timer:sw22p.Transx=0:sw22.TimerEnabled=0:End Sub
Sub sw23_Timer:sw23p.Transx=0:sw23.TimerEnabled=0:End Sub
Sub sw24_Timer:sw24p.Transx=0:sw24.TimerEnabled=0:End Sub

Sub sw27_Timer:sw27p.Transx=0:sw27.TimerEnabled=0:End Sub
Sub sw28_Timer:sw28p.Transx=0:sw28.TimerEnabled=0:End Sub
Sub sw29_Timer:sw29p.Transx=0:sw29.TimerEnabled=0:End Sub


'Rollovers
Sub sw14_Hit:Controller.Switch(14) = 1:PlaySound "rollover":End Sub
Sub sw14_UnHit:Controller.Switch(14) = 0:End Sub
Sub sw15_Hit:Controller.Switch(15) = 1:PlaySound "rollover":End Sub
Sub sw15_UnHit:Controller.Switch(15) = 0:End Sub
Sub sw35a_Hit:Controller.Switch(35) = 1:PlaySound "rollover":DOF 101, 1:End Sub
Sub sw35a_UnHit:Controller.Switch(35) = 0:DOF 101, 0:End Sub
Sub sw35b_Hit:Controller.Switch(35) = 1:PlaySound "rollover":DOF 102, 1:End Sub
Sub sw35b_UnHit:Controller.Switch(35) = 0:DOF 102, 0:End Sub
Sub sw17_Hit:Controller.Switch(17) = 1:PlaySound "rollover":End Sub
Sub sw17_UnHit:Controller.Switch(17) = 0:End Sub
Sub sw18_Hit:Controller.Switch(18) = 1:PlaySound "rollover":End Sub
Sub sw18_UnHit:Controller.Switch(18) = 0:End Sub

'Rollover Buttons
Sub sw25a_Hit:Controller.Switch(25) = 1:PlaySound "rollover":End Sub
Sub sw25a_UnHit:Controller.Switch(25) = 0:End Sub
Sub sw25b_Hit:Controller.Switch(25) = 1:PlaySound "rollover":End Sub
Sub sw25b_UnHit:Controller.Switch(25) = 0:End Sub
Sub sw26_Hit:Controller.Switch(26) = 1:PlaySound "rollover":End Sub
Sub sw26_UnHit:Controller.Switch(26) = 0:End Sub

sub RealTime_timer
	GateHeavyT1.RotZ = -(Gate.currentangle)
	GateHeavyT2.RotZ = -(Gate1.currentangle)
    DiverterP.RotY= Diverter.currentangle
end sub

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

Sub RandomSoundRubber()
	Select Case Int(Rnd*3)+1
		Case 1 : PlaySound "rubber_hit_1", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
		Case 2 : PlaySound "rubber_hit_2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
		Case 3 : PlaySound "rubber_hit_3", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
	End Select
End Sub



' *********************************************************************
'                      Supporting Ball & Sound Functions
' *********************************************************************

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
    Vol = Csng((BallVel(ball)*0.4) + 6)
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

Const tnob = 6 ' total number of balls
ReDim rolling(6)
Dim collision(6)
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


'************************************
' What you need to add to your table
'************************************

' a timer called CollisionTimer. With a fast interval, like from 1 to 10
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

' The Double For loop: This is a double cycle used to check the collision between a ball and the other ones.
' If you look at the parameters of both cycles, you’ll notice they are designed to avoid checking 
' collision twice. For example, I will never check collision between ball 2 and ball 1, 
' because I already checked collision between ball 1 and 2. So, if we have 4 balls, 
' the collision checks will be: ball 1 with 2, 1 with 3, 1 with 4, 2 with 3, 2 with 4 and 3 with 4.

' Sum first the radius of both balls, and if the height between them is higher then do not calculate anything more,
' since the balls are on different heights so they can't collide.

' The next 3 lines calculates distance between xth and yth balls with the Pytagorean theorem,

' The first "If": Checking if the distance between the two balls is less than the sum of the radius of both balls, 
' and both balls are not already colliding.

' Why are we checking if balls are already in collision? 
' Because we do not want the sound repeting when two balls are resting closed to each other.

' Set the collision property of both balls to True, and we assign the number of the ball colliding

' Play the collide sound of your choice using the VOL, PAN and PITCH functions

' Last lines: If the distance between 2 balls is more than the radius of a ball,
' then there is no collision and then set the collision property of the ball to False (-1).


Sub Pins_Hit (idx)
	PlaySound "pinhit_low", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0
End Sub

Sub Targets_Hit (idx)
	PlaySound SoundFX("target"), 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0
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

'Bally Six million dollar man
'From Inkochnito's addition in the previous pinmame table
Sub editDips
	Dim vpmDips : Set vpmDips = New cvpmDips
	With vpmDips
		.AddForm 700,400,"The Six Million Dollar Man - DIP switches"
		.AddChk 7,10,180,Array("Match feature",&H00100000)'dip 21
		.AddChk 205,10,115,Array("Credits display",&H00080000)'dip 20
		.AddFrame 2,30,190,"Maximum credits",&H00070000,Array("10 credits",&H00010000,"15 credits",&H00020000,"25 credits",&H00040000,"40 credits",&H00070000)'dip 17&18&19
		.AddFrame 2,106,190,"Sound features",&H80000080,Array("chime effects",&H80000000,"chime and tunes",0,"noise",&H00000080,"noises and tunes",&H80000080)'dip 8&32
		.AddFrame 2,184,190,"High score feature",&H00006000,Array("no award",0,"extra ball",&H00004000,"replay",&H00006000)'dip 14&15
		.AddFrame 2,248,190,"50,000 special",&H10000000,Array("3X on 1 ball = special",0,"2X on 1 ball = special",&H10000000)'dip 29
		.AddFrame 205,30,190,"High score to date",&H00000060,Array("no award",0,"1 credit",&H00000020,"2 credits",&H00000040,"3 credits",&H00000060)'dip 6&7
		.AddFrame 205,106,190,"Balls per game", 32768,Array("3 balls",0,"5 balls", 32768)'dip 16
		.AddFrame 205,152,190,"Outlane special",&H00200000,Array("alternates from side to side",0,"both lanes lite for special",&H00200000)'dip 22
		.AddFrame 205,200,190,"Saucer adjustment",&H00400000,Array ("saucer starts at 3,000",0,"saucer starts at 5,000",&H00400000)'dip 23
		.AddFrame 205,248,190,"Upper and lower O target",&H00800000,Array("are not tied",0,"are tied together",&H00800000)'dip 24
		'dip 30&31 not used
		.AddLabel 50,310,300,20,"After hitting OK, press F3 to reset game with new settings."
		.ViewDips
	End With
End Sub
Set vpmShowDips = GetRef("editDips")
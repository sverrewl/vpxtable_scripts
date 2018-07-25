'############################################################################################
'############################################################################################
'#######                                                                             ########
'#######                                  NUGENT                                     ########
'#######                               (Stern 1978)                                  ########
'#######                                                                             ########
'############################################################################################
'############################################################################################
' Version 1.0.1 WED21 2017

' Many thanks go out to:
' 32Assassin for script help
' Allknowing2012 for script help
' Hauntfreaks for the Buttons and for his knowledge!!
' GTXJoe for his primitive collection
' Flupper for the physics starting point
' NFozzy for even more script help and adding his "FastFlips"

' Thalamus 2018-07-24
' Table has its own "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Changed UseSolenoids=1 to 2
' No special SSF tweaks yet.

Option Explicit
Randomize


On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0


Const cGameName="nugent",UseSolenoids=2,UseLamps=1,UseGI=0,SSolenoidOn="SolOn",SSolenoidOff="SolOff", SCoin="fx_coin"
'Const ballmass = 1.7
Const ballmass = 2.0
LoadVPM "01120100","Bally.vbs",3.10



'*************************************************************
'Solenoid Call backs
'**********************************************************************************************************

SolCallback(10)="DropLeft"
SolCallback(5)="dtR.SolDropUp"
SolCallback(6)="vpmSolSound SoundFX(""Knocker"",DOFKnocker),"
SolCallback(7)="bsTrough.SolOut"
SolCallback(9)="bsSaucer.SolOut"
'SolCallback(19)="vpmNudge.SolGameOn"
SolCallback(19)="FastFlips.TiltSol"
SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub DropLeft(Enabled)
    If Enabled Then
        dtL.DropSol_On
        Controller.Switch(13)=0
        Controller.Switch(37)=0
    End If
End Sub


Sub SolLFlipper(Enabled)
     If Enabled Then
         PlaySoundAt SoundFX("fx_FlipperupLeft",DOFContactors), LeftFlipper:LeftFlipper.RotateToEnd
     Else
         PlaySoundAt SoundFX("fx_FlipperdownLeft",DOFContactors), LeftFlipper:LeftFlipper.RotateToStart
     End If
  End Sub

Sub SolRFlipper(Enabled)
     If Enabled Then
         PlaySoundAt SoundFX("fx_FlipperupRight",DOFContactors), RightFlipper:RightFlipper.RotateToEnd: RightFlipperUpper.RotateToEnd
     Else
         PlaySoundAt SoundFX("fx_FlipperdownRight",DOFContactors), RightFlipper:RightFlipper.RotateToStart: RightFlipperUpper.RotateToStart
     End If
End Sub




Sub Table1_KeyDown(ByVal KeyCode)
    If KeyDownHandler(keycode) Then Exit Sub
    If keycode = PlungerKey Then Plunger.Pullback:PlaySoundAt "plungerpull", Plunger
    if KeyCode = LeftTiltKey Then Nudge 90, 4
    if KeyCode = RightTiltKey Then Nudge 270, 4
    if KeyCode = CenterTiltKey Then Nudge 0, 4
    If KeyCode = LeftFlipperKey then FastFlips.FlipL True :  FastFlips.FlipUL True
    If KeyCode = RightFlipperKey then FastFlips.FlipR True :  FastFlips.FlipUR True
        'If keycode = LeftFlipperKey Then LeftFlipper.RotateToEnd: PlaySoundAt SoundFX("fx_flipperupLeft",DOFContactors), LeftFlipper
        'If keycode = RightFlipperKey Then RightFlipper.RotateToEnd: PlaySoundAt SoundFX("fx_flipperupRight",DOFContactors), RightFlipper
        'If keycode = RightFlipperKey Then RightFlipperUpper.RotateToEnd
End Sub

Sub Table1_KeyUp(ByVal KeyCode)
    If KeyUpHandler(keycode) Then Exit Sub
    If keycode = PlungerKey Then Plunger.Fire:PlaySoundAt "plunger", Plunger
    If KeyCode = LeftFlipperKey then FastFlips.FlipL False :  FastFlips.FlipUL False
    If KeyCode = RightFlipperKey then FastFlips.FlipR False :  FastFlips.FlipUR False
        'If keycode = LeftFlipperKey Then LeftFlipper.RotateToStart: PlaySound SoundFX("fx_flipperdownLeft",DOFContactors)
        'If keycode = RightFlipperKey Then RightFlipper.RotateToStart: PlaySound SoundFX("fx_flipperdownRight",DOFContactors)
        'If keycode = RightFlipperKey Then RightFlipperUpper.RotateToStart
End Sub

Sub SolKnocker(Enabled)
    If Enabled Then PlaySound SoundFX("Knocker",DOFKnocker)
End Sub


'***********************************************************************

'Initiate Table
'**********************************************************************************************************

Dim bsTrough, dtL, dtR, bsSaucer, FastFlips

Sub Table1_Init
    vpmInit Me
    On Error Resume Next
        With Controller
        .GameName = cGameName
        If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
        .SplashInfoLine = "Nugent -- Stern 1978"&chr(13)&"ROCK-n-ROLL"
        .HandleMechanics=0
        .HandleKeyboard=0
        .ShowDMDOnly=1
        .ShowFrame=0
        .ShowTitle=0
            .hidden = not table1.showdt
         On Error Resume Next
         .Run GetPlayerHWnd
         If Err Then MsgBox Err.Description
         On Error Goto 0
     End With
     On Error Goto 0

    PinmameTimer.Interval=PinMameInterval
    PinmameTimer.Enabled=1
    vpmNudge.TiltSwitch=7
    vpmNudge.Sensitivity=2
    vpmNudge.TiltObj=Array(Bumper1,Bumper2,Bumper3)

    Set FastFlips = new cFastFlips
    with FastFlips
        .CallBackL = "SolLflipper"  'Point these to flipper subs
        .CallBackR = "SolRflipper"  '...
    '   .CallBackUR = "SolURflipper"'...
        .TiltObjects = True 'Optional, if True calls vpmnudge.solgameon automatically
    '   .DebugOn = False        'Debug, always-on flippers. Call FastFlips.DebugOn True or False in debugger to enable/disable.
    end with


'************TROUGH************************
    Set bsTrough=New cvpmBallstack
        bsTrough.InitNoTrough BallRelease,8,120,4
        bsTrough.InitExitSnd SoundFX("fx_ballrelease",DOFContactors), SoundFX("Solenoid",DOFContactors)

'************TOP KICKER********************
    Set bsSaucer=New cvpmBallStack
        bsSaucer.InitSaucer Kicker1,24,90,20
        bsSaucer.InitExitSnd SoundFX("Popper_Ball",DOFContactors), SoundFX("Solenoid",DOFContactors)
        bsSaucer.KickAngleVar=30

'************DROP TARGETS*******************

    Set dtR = New cvpmDropTarget
        dtR.InitDrop Array(DT4,DT5,DT6),Array(30,22,14)
        dtR.InitSnd SoundFX("fx_DTDrop1",DOFDropTargets),SoundFX("fx_DTReset",DOFContactors)

    Set dtL = New cvpmDropTarget
        dtL.InitDrop Array(DT1,DT2,DT3),Array(5,21,29)
        dtL.InitSnd SoundFX("fx_DTDrop1",DOFDropTargets),SoundFX("fx_DTReset",DOFContactors)

 vpmMapLights AllLights
 Set Lights(61)  = GIlamps 'ball in play

 End Sub

'GI uselamps workaround (lamp spoof)
dim GIlamps : set GIlamps = New GIcatcherobject
Class GIcatcherObject   'object that disguises itself as a light. (UseLamps workaround for System80 GI circuit)
    Public Property Let State(input)
        dim x
        if input = 1 then 'If GI switch is engaged, turn off GI.
            for each x in gi : x.state = 0 : next
        elseif input = 0 then
            for each x in gi : x.state = 1 : next
        end if
    End Property
End Class



'*************DROP TARGETS********

'LEFT BANK
Sub DT1_Dropped:dtL.Hit 1:PlaysoundAt "fx_DTDrop",DT1:Controller.Switch(13)=1:End Sub
Sub DT2_Dropped:dtL.Hit 2:PlaysoundAt "fx_DTDrop",DT2:End Sub
Sub DT3_Dropped:dtL.Hit 3:PlaysoundAt "fx_DTDrop",DT3:Controller.Switch(37)=1:End Sub

'RIGHT BANK
Sub DT4_Dropped:dtR.Hit 1:PlaysoundAt "fx_DTDrop",DT4:End Sub
Sub DT5_Dropped:dtR.Hit 2:PlaysoundAt "fx_DTDrop",DT5:End Sub
Sub DT6_Dropped:dtR.Hit 3:PlaysoundAt "fx_DTDrop",DT6:End Sub

'**********************************************************************************************************

'***************DRAIN AND KICKER*******************

Sub Drain_Hit:bsTrough.addball me : PlaySoundAt "fx_DrainShort",Drain: LightInlaneLeft.state = 0:LightInlaneRight.state = 0: End Sub
Sub Kicker1_Hit:bsSaucer.AddBall 0 : PlaySoundAt "KickerEnter",kicker1: End Sub



'**************************************************
'SWITCHES
'**************************************************

Sub SW40A_hit:vpmtimer.pulsesw 40:LightInlaneLeft.state = 0:LightInlaneRight.state = 0:End Sub
Sub SW40A_unhit:Controller.Switch(40)=0:End Sub

Sub SW40B_hit:vpmtimer.pulsesw 40:LightInlaneLeft.state = 0:LightInlaneRight.state = 0:End Sub
Sub SW40B_unhit:Controller.Switch(40)=0:End Sub

Sub SW3A_hit:vpmtimer.pulsesw 3:End Sub
Sub SW3A_unhit:Controller.Switch(3)=0:End Sub

Sub SW3B_hit:vpmtimer.pulsesw 3:End Sub
Sub SW3B_unhit:Controller.Switch(3)=0:End Sub

Sub SW1B_hit:vpmtimer.pulsesw 1:End Sub
Sub SW1B_unhit:Controller.Switch(1)=0:End Sub

Sub SW1A_hit:vpmtimer.pulsesw 1:End Sub
Sub SW1A_unhit:Controller.Switch(1)=0:End Sub

Sub SW4A_hit:vpmtimer.pulsesw 4:End Sub
Sub SW4A_unhit:Controller.Switch(4)=0:End Sub

Sub SW4B_hit:vpmtimer.pulsesw 4:End Sub
Sub SW4B_unhit:Controller.Switch(4)=0:End Sub

Sub SW2_hit:vpmtimer.pulsesw 2:End Sub
Sub SW2_unhit:Controller.Switch(2)=0:End Sub

Sub SW17A_hit:vpmtimer.pulsesw 17:End Sub
Sub SW17A_unhit:Controller.Switch(17)=0:End Sub

Sub SW17B_hit:vpmtimer.pulsesw 17:End Sub
Sub SW17B_unhit:Controller.Switch(17)=0:End Sub

Sub SW8_hit:vpmtimer.pulsesw 8:End Sub
Sub SW8_unhit:Controller.Switch(8)=0:End Sub

Sub SW19_hit:vpmtimer.pulsesw 19:End Sub

Sub SW18_hit:vpmtimer.pulsesw 18:End Sub



'**************SPINNER************
Sub Spinner1_Spin():vpmTimer.PulseSw 15:PlaySoundAt "fx_spinner",Spinner1 :End Sub


'**************TARGET*************
Sub Target1_hit:vpmTimer.PulseSw 12:PlaySoundAt "Target", Target1
                LightInlaneRight.state = 1
                LightInlaneLeft.state = 1
End Sub


'*************BUMPERS*************

Sub Bumper1_hit:vpmTimer.PulseSw 35:PlaySoundAt "Fx_Bumper1", Bumper1:End Sub

Sub Bumper2_hit:vpmTimer.PulseSw 34:PlaySoundAt "Fx_Bumper2", Bumper2:End Sub

Sub Bumper3_hit:vpmTimer.PulseSw 33:PlaySoundAt "Fx_Bumper3", Bumper3:End Sub


 'added by Inkochnito
Sub editDips
	Dim vpmDips:Set vpmDips=New cvpmDips
	With vpmDips
		.AddForm 700,400,"Nugent - DIP switches"
		.AddChk 2,10,180,Array("Match feature",&H00100000)'dip 21
		.AddChk 205,10,115,Array("Credits display",&H00080000)'dip 20
		.AddFrame 2,30,190,"Maximum credits",&H00070000,Array("5 credits",0,"10 credits",&H00010000,"15 credits",&H00020000,"20 credits",&H00030000,"25 credits",&H00040000,"30 credits",&H00050000,"35 credits",&H00060000,"40 credits",&H00070000)'dip 17&18&19
		.AddFrame 2,160,190,"High game to date award",49152,Array("points",0,"1 free game",&H00004000,"2 free games",32768,"3 free games",49152)'dip 15&16
		.AddFrame 2,235,190,"Special award",&HC0000000,Array("100,000 points",0,"free ball",&H40000000,"free game",&H80000000,"free ball and free game",&HC0000000)'dip 31&32
		.AddFrame 2,310,190,"High score feature",&H00000020,Array("extra ball",0,"replay",&H00000020)'dip 6
		.AddFrame 205,30,190,"Balls per game",&H00000040,Array("3 balls",0,"5 balls",&H00000040)'dip 7
		.AddFrame 205,76,190,"Return lane lamps",&H00002000,Array("goes off after being made",0,"stay on for entire ball",&H00002000)'dip 14
		.AddFrame 205,122,190,"Special limit 3 bank drop targets",&H00200000,Array("no limit",0,"one special per ball",&H00200000)'dip 22
		.AddFrame 205,169,190,"Electronic sound",&H00400000,Array("electronic chimes",0,"computer type sounds",&H00400000)'dip 23
		.AddFrame 205,216,190,"Bonus countdown",&H00800000,Array("multiple steps",0,"1,000 per step",&H00800000)'dip 24
		.AddFrame 205,263,190,"Extra ball",&H02000000,Array("no extra ball (bypass)",0,"award extra ball",&H02000000)'dip 26
		.AddFrame 205,310,190,"Melody option",&H00000080,Array("2 tones only",0,"full melody",&H00000080)'dip 8
		.AddLabel 50,382,300,20,"After hitting OK, press F3 to reset game with new settings."
		.ViewDips
	End With
End Sub
Set vpmShowDips=GetRef("editDips")


'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep

Sub RightSlingShot_Slingshot
    vpmTimer.PulseSw(39)
    PlaySoundat "SlingshotRight", Peg1
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
    vpmTimer.PulseSw(38)
    PlaySoundAt "SlingshotLeft",Peg6
    LSling.Visible = 0
    LSling1.Visible = 1
    sling2.TransZ = -20
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

Const tnob = 1 ' total number of balls
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



''*****************************************
''      ninuzzu's   FLIPPER SHADOWS
''*****************************************
'
sub FlipperTimer_Timer()
    FlipperLSh.RotZ = LeftFlipper.currentangle
    FlipperRSh.RotZ = RightFlipper.currentangle
    FlipperRUSh.RotZ = RightFlipperUpper.currentangle
End Sub


'*****************************************
'   ninuzzu's   BALL SHADOW
'*****************************************
Dim BallShadow
BallShadow = Array (BallShadow1)

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


Sub Pins_Hit (idx)
    PlaySound "fx_PinHit", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
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

Sub Switches_Hit (idx)
    PlaySound "fx_sensor", 0, 1, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Rollovers_Hit (idx)
    PlaySound "Rollover", 0, 0.50, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Gates_Hit (idx)
    PlaySound "GateFlap", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
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

Sub RightFlipperUpper_Collide(parm)
    RandomSoundFlipper()
End Sub


Sub RandomSoundFlipper()
    Select Case Int(Rnd*3)+1
        Case 1 : PlaySound "flip_hit_1", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
        Case 2 : PlaySound "flip_hit_2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
        Case 3 : PlaySound "flip_hit_3", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    End Select
End Sub

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

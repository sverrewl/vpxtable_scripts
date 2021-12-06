                                                                 'VISUAL PINBALL 10
                                                               'TAITO GEMINI 2000 1982
                                                              'Terminada em  18/10/2017
                                                       'Emulada com a colaboração de Doxpinball, Roataguile e Tioitalo

' Thalamus 2018-07-24
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Table has a unusual ball rolling routine
' No special SSF tweaks yet.


Option Explicit

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Const cGameName="gemini",UseSolenoids=1,UseLamps=1
Const SSolenoidOn="solon",SSolenoidOff="soloff",SFlipperOn="flipperup",SFlipperOff="flipperdown",SCoin="coin"

LoadVPM "01200100","Taito.VBS",3.1
Dim DesktopMode: DesktopMode = Table1.ShowDT

If DesktopMode = True Then 'Show Desktop components
Ramp16.visible=1
Ramp15.visible=1
Primitive52.visible=1
Else
Ramp16.visible=0
Ramp15.visible=0
Primitive52.visible=0
End if

'*************************************************************
'Solenoid Call backs
'**********************************************************************************************************

SolCallback(2)    = "bsTrough.SolOut"
SolCallback(3)    = "vpmSolAutoPlunger Plunger1,1,"
SolCallback(4)    = "dtC.SolDropUp"
SolCallback(5)    = "dtR.SolDropUp"
SolCallback(7)    = "dtR.SolHit 1,"
SolCallback(8)    = "dtR.SolHit 2,"
SolCallback(9)    = "dtR.SolHit 3,"
SolCallback(10)   = "dtR.SolHit 4,"
SolCallback(11)   = "dtT.SolDropUp"
SolCallback(18)   = "vpmNudge.SolGameOn"

SolCallback(sLRFlipper)     = "SolRFlipper"         ' Right Flipper
SolCallback(sLLFlipper)     = "SolLFlipper"         ' Left Flipper

Sub SolLFlipper(Enabled)
     If Enabled Then
         PlaySound SoundFX("fx_Flipperup",DOFContactors):LeftFlipper.RotateToEnd
     Else
         PlaySound SoundFX("fx_Flipperdown",DOFContactors):LeftFlipper.RotateToStart
     End If
  End Sub

Sub SolRFlipper(Enabled)
     If Enabled Then
         PlaySound SoundFX("fx_Flipperup",DOFContactors):RightFlipper.RotateToEnd
     Else
         PlaySound SoundFX("fx_Flipperdown",DOFContactors):RightFlipper.RotateToStart
     End If
End Sub

'**********************************************************************************************************

'Solenoid Controlled toys
'**********************************************************************************************************

'*****GI Lights On
dim xx
For each xx in GI:xx.State = 1: Next

'**********************************************************************************************************

'Initiate Table
'**********************************************************************************************************

Dim bsTrough, dtC, dtR, dtT, cbCaptive

Sub Table1_Init
  On Error Resume Next
    With Controller
      .GameName=cGameName
      If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
      .SplashInfoLine="Taito Gemini 2000"
      .HandleMechanics=0
      .ShowDMDOnly=1
      .ShowFrame=0
      .ShowTitle=0
      .Run
            .Hidden=1
      If Err Then MsgBox Err.Description
    End With
  On Error Goto 0

  vpmNudge.TiltSwitch=30
  vpmNudge.Sensitivity=5
  vpmNudge.TiltObj=Array(Bumper72e,Bumper72d)

  Set bsTrough=New cvpmBallStack
  bsTrough.InitSw 0,51,41,31,11,0,0,0
  bsTrough.InitKick BallRelease,60,10
  bsTrough.InitExitSnd SoundFX("ballrelease",DOFContactors), SoundFX("Solenoid",DOFContactors)
  bsTrough.Balls=4

  Set cbCaptive=New cvpmCaptiveBall
  cbCaptive.InitCaptive Trigger1,Block,kicker1,0
  cbCaptive.Start
  cbCaptive.ForceTrans=.99
  cbCaptive.MinForce=15.5
  cbCaptive.CreateEvents "cbCaptive"

  Set dtC=New cvpmDropTarget
    dtC.InitDrop Array(D43,D53,D63,D73),Nothing
    dtC.InitSnd SoundFX("DTDrop",DOFContactors),SoundFX("DTReset",DOFContactors)
    dtC.CreateEvents "dtC"

  Set dtR=New cvpmDropTarget
  dtR.InitDrop Array(D44,D54,D64,D74),Nothing
  dtR.InitSnd SoundFX("DTDrop",DOFContactors),SoundFX("DTReset",DOFContactors)
  dtR.CreateEvents "dtR"

  Set dtT=New cvpmDropTarget
  dtT.InitDrop Array(D5,D15,D25,D35),Nothing
  dtT.InitSnd SoundFX("DTDrop",DOFContactors),SoundFX("DTReset",DOFContactors)
  dtT.CreateEvents "dtT"

  vpmCreateEvents Contatos
  vpmMapLights Lampadas
  Plunger1.PullBack

End Sub
'**********************************************************************************************************
'Plunger code
'**********************************************************************************************************
Sub Table1_KeyDown(ByVal KeyCode)
  If KeyCode=RightFlipperKey Then Controller.Switch(71)=1
  If vpmKeyDown(KeyCode) Then Exit Sub
  If KeyCode=PlungerKey Then Controller.Switch(65)=1
End Sub

Sub Table1_KeyUp(ByVal KeyCode)
  If KeyCode=RightFlipperKey Then Controller.Switch(71)=0
  If vpmKeyUp(KeyCode) Then Exit Sub
  If KeyCode=PlungerKey Then Controller.Switch(65)=0
End Sub
'**********************************************************************************************************

 ' Drain hole and kickers
Sub Drain_Hit:bsTrough.addball me : playsound"drain" : End Sub

 'Scoring Rubber
Sub Slingshot2a_Slingshot:vpmTimer.PulseSw 2 : playsound"flip_hit_3" : End Sub
Sub Slingshot2b_Slingshot:vpmTimer.PulseSw 2 : playsound"flip_hit_3" : End Sub
Sub Slingshot2c_Slingshot:vpmTimer.PulseSw 2 : playsound"flip_hit_3" : End Sub
Sub Slingshot2d_Slingshot:vpmTimer.PulseSw 2 : playsound"flip_hit_3" : End Sub

'Bumpers
Sub Bumper72e_Hit : vpmTimer.PulseSw(72) : playsound SoundFX("fx_bumper1",DOFContactors): End Sub
Sub Bumper72d_Hit : vpmTimer.PulseSw(72) : playsound SoundFX("fx_bumper1",DOFContactors): End Sub

 'Stand Up Targets
Sub Alvo12_Hit:vpmTimer.PulseSw 12 : End Sub
Sub Alvo22_Hit:vpmTimer.PulseSw 22 : End Sub
Sub Alvo32_Hit:vpmTimer.PulseSw 32 : End Sub
Sub Alvo42_Hit:vpmTimer.PulseSw 42 : End Sub
Sub Alvo52_Hit:vpmTimer.PulseSw 52 : End Sub
Sub Alvo55_Hit:vpmTimer.PulseSw 55 : End Sub

Sub Alvo75_Hit:vpmTimer.PulseSw 75 : End Sub



'Display do placar
Dim Digits(27)
Digits(0)=Array(Light5,Light6,Light7,Light8,Light9,Light10,Light11)
Digits(1)=Array(Light12,Light13,Light14,Light15,Light16,Light17,Light18)
Digits(2)=Array(Light19,Light20,Light21,Light22,Light23,Light24,Light25)
Digits(3)=Array(Light26,Light27,Light28,Light29,Light30,Light31,Light32)
Digits(4)=Array(Light33,Light34,Light35,Light36,Light37,Light38,Light39)
Digits(5)=Array(Light40,Light41,Light42,Light43,Light44,Light45,Light46)
Digits(6)=Array(Light47,Light48,Light49,Light50,Light51,Light52,Light53)
Digits(7)=Array(Light54,Light55,Light56,Light57,Light58,Light59,Light60)
Digits(8)=Array(Light61,Light62,Light63,Light64,Light65,Light66,Light67)
Digits(9)=Array(Light68,Light69,Light70,Light71,Light72,Light73,Light74)
Digits(10)=Array(Light75,Light76,Light77,Light78,Light79,Light80,Light81)
Digits(11)=Array(Light82,Light83,Light84,Light85,Light86,Light87,Light88)
Digits(12)=Array(Light89,Light90,Light91,Light92,Light93,Light94,Light95)
Digits(13)=Array(Light96,Light97,Light98,Light99,Light100,Light101,Light102)
Digits(14)=Array(Light103,Light104,Light105,Light106,Light107,Light108,Light109)
Digits(15)=Array(Light110,Light111,Light112,Light113,Light114,Light115,Light116)
Digits(16)=Array(Light117,Light118,Light119,Light120,Light121,Light122,Light123)
Digits(17)=Array(Light124,Light125,Light126,Light127,Light128,Light129,Light130)
Digits(18)=Array(Light131,Light132,Light133,Light134,Light135,Light136,Light137)
Digits(19)=Array(Light138,Light139,Light140,Light141,Light142,Light143,Light144)
Digits(20)=Array(Light145,Light146,Light147,Light148,Light149,Light150,Light151)
Digits(21)=Array(Light152,Light153,Light154,Light155,Light156,Light157,Light158)
Digits(22)=Array(Light159,Light160,Light161,Light162,Light163,Light164,Light165)
Digits(23)=Array(Light166,Light167,Light168,Light169,Light170,Light171,Light172)
Digits(24)=Array(Light173,Light174,Light175,Light176,Light177,Light178,Light179)
Digits(25)=Array(Light180,Light181,Light182,Light183,Light184,Light185,Light186)
Digits(26)=Array(Light187,Light188,Light189,Light190,Light191,Light192,Light193)
Digits(27)=Array(Light194,Light195,Light196,Light197,Light198,Light199,Light200)

Sub DisplayTimer_Timer
If DesktopMode = True Then 'Show Desktop components
  Dim ChgLED,ii,num,chg,stat,obj
  ChgLED=Controller.ChangedLEDs(&Hffffffff,&Hffffffff)
  If Not IsEmpty(ChgLED) Then
    For ii=0 To UBound(chgLED)
    num=chgLED(ii,0):chg=chgLED(ii,1):stat=chgLED(ii,2)
      For Each obj In Digits(num)
        If chg And 1 Then obj.State=stat And 1
        chg=chg\2:stat=stat\2
      Next
    Next
  End If
End if
End Sub



'**********************************************************************************************************
'**********************************************************************************************************
' Start of VPX functions
'**********************************************************************************************************
'**********************************************************************************************************


'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep

Sub Slingshot45d_Slingshot
  vpmTimer.PulseSw 45
    PlaySound SoundFX("right_slingshot",DOFContactors), 0, 1, 0.05, 0.05
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.TransZ = -20
    RStep = 0
    Slingshot45d.TimerEnabled = 1
End Sub

Sub SlingShot45d_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.TransZ = -10
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.TransZ = 0:Slingshot45d.TimerEnabled = 0:
    End Select
    RStep = RStep + 1
End Sub

Sub Slingshot45e_Slingshot
  vpmTimer.PulseSw 45
    PlaySound SoundFX("left_slingshot",DOFContactors),0,1,-0.05,0.05
    LSling.Visible = 0
    LSling1.Visible = 1
    sling2.TransZ = -20
    LStep = 0
    Slingshot45e.TimerEnabled = 1
End Sub

Sub Slingshot45e_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.TransZ = -10
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.TransZ = 0:Slingshot45e.TimerEnabled = 0:
    End Select
    LStep = LStep + 1
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
  PlaySound "pinhit_low", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Targets_Hit (idx)
  PlaySound "target", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Metals_Thin_Hit (idx)
  PlaySound "metalhit_thin", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Metals_Medium_Hit (idx)
  PlaySound "metalhit_medium", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Metals2_Hit (idx)
  PlaySound "metalhit2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Gates_Hit (idx)
  PlaySound "gate4", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Spinner_Spin
  PlaySound "fx_spinner",0,.25,0,0.25
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
  PlaySound sound, 1, Vol, Pan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

'Set all as per ball position & speed, but Vol Multiplier may be used eg; PlaySoundAtBallVol "sound",3

Sub PlaySoundAtBallVol(sound, VolMult)
  PlaySound sound, 0, Vol(ActiveBall) * VolMult, Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 1, AudioFade(ActiveBall)
End Sub

'Set position as bumperX and Vol manually.

Sub PlaySoundAtBumperVol(sound, tableobj, Vol)
  PlaySound sound, 1, Vol, Pan(tableobj), 0,0,1, 1, AudioFade(tableobj)
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

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "table1" is the name of the table
    Dim tmp
    tmp = ball.x * 2 / table1.width-1
    If tmp > 0 Then
        Pan = Csng(tmp ^10)
    Else
        Pan = Csng(-((- tmp) ^10) )
    End If
End Function

Function AudioFade(ball) ' Can this be together with the above function ?
  Dim tmp
  tmp = ball.y * 2 / Table1.height-1
  If tmp > 0 Then
    AudioFade = Csng(tmp ^10)
  Else
    AudioFade = Csng(-((- tmp) ^10) )
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
'    JP's VP10 Collision & Rolling Sounds
'*****************************************

Const tnob = 8 ' total number of balls
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



' Thalamus : Exit in a clean and proper way
Sub Table1_exit()
  Controller.Pause = False
  Controller.Stop
End Sub


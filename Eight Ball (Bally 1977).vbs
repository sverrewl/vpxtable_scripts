Option Explicit
Randomize

' Thalamus 2018-07-20
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Changed UseSolenoids=1 to 2

' Thalamus 2019 August : Improved directional sounds
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

Const cGameName="eightbll",UseSolenoids=2,UseLamps=1,UseGI=0,SSolenoidOn="SolOn",SSolenoidOff="SolOff", SCoin="coin"

LoadVPM "01320000", "Bally.VBS", 3.1
Dim DesktopMode: DesktopMode = Table1.ShowDT

If DesktopMode = True Then 'Show Desktop components
Ramp16.visible=1
Ramp15.visible=1
Primitive13.visible=1
Else
Ramp16.visible=0
Ramp15.visible=0
Primitive13.visible=0
End if

'*************************************************************
'Solenoid Call backs
'**********************************************************************************************************
SolCallback(2)    =  "vpmSolSound SoundFX(""Bell10"",DOFKnocker),"
SolCallback(3)    =  "vpmSolSound SoundFX(""Bell100"",DOFKnocker),"
SolCallback(4)    =  "vpmSolSound SoundFX(""Bell1000"",DOFKnocker),"
SolCallback(5)    =  "vpmSolSound SoundFX(""Bell10000"",DOFKnocker),"
SolCallback(6)    =  "vpmSolSound SoundFX(""Knocker"",DOFKnocker),"
SolCallback(7)    =  "bsTrough.SolOut"
SolCallback(8)    =  "SolKickBack"
SolCallback(19)   =  "vpmNudge.SolGameOn"

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
     If Enabled Then
         PlaySoundAtVol SoundFX("fx_Flipperup",DOFContactors), LeftFlipper, 1:LeftFlipper.RotateToEnd
     Else
         PlaySoundAtVol SoundFX("fx_Flipperdown",DOFContactors), LeftFlipper, 1:LeftFlipper.RotateToStart
     End If
  End Sub

Sub SolRFlipper(Enabled)
     If Enabled Then
         PlaySoundAtVol SoundFX("fx_Flipperup",DOFContactors), RightFlipper, 1:RightFlipper.RotateToEnd
     Else
         PlaySoundAtVol SoundFX("fx_Flipperdown",DOFContactors), RightFlipper, 1:RightFlipper.RotateToStart
     End If
End Sub
'**********************************************************************************************************

'Solenoid Controlled toys
'**********************************************************************************************************

'*****GI Lights On
dim xx
For each xx in GI:xx.State = 1: Next

Sub SolKickBack(enabled)
  if enabled then
    KickBack.Fire
    playsound SoundFX("Popper",DOFContactors)
  else
    KickBack.PullBack
  end if
End Sub

'**********************************************************************************************************

'Initiate Table
'**********************************************************************************************************

Dim bsTrough

Sub Table1_Init
  vpmInit Me
  On Error Resume Next
    With Controller
    .GameName = cGameName
    If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
    .SplashInfoLine = "Eight Ball Bally"&chr(13)&"You Suck"
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
  PinMAMETimer.Enabled = True

  vpmNudge.TiltSwitch = 7
  vpmNudge.Sensitivity = 1
  vpmNudge.TiltObj = Array(Bumper1,Bumper2,Bumper3,LeftSlingshot,RightSlingShot)

  Set bsTrough = New cvpmBallStack ' Trough handler
    bsTrough.InitSw 0,8,0,0,0,0,0,0
    bsTrough.InitKick BallRelease, 50, 4
    bsTrough.InitExitSnd SoundFX("ballrelease",DOFContactors), SoundFX("Solenoid",DOFContactors)
    bsTrough.Balls = 1

  KickBack.Pullback

End Sub

'**********************************************************************************************************
'Plunger code
'**********************************************************************************************************

Sub Table1_KeyDown(ByVal KeyCode)
  If KeyDownHandler(keycode) Then Exit Sub
  If keycode = PlungerKey Then Plunger.Pullback:playsoundAtVol"plungerpull", Plunger, 1
End Sub

Sub Table1_KeyUp(ByVal KeyCode)
  If KeyUpHandler(keycode) Then Exit Sub
  If keycode = PlungerKey Then Plunger.Fire:PlaySoundAtVol"plunger", Plunger, 1
End Sub

'**********************************************************************************************************

 ' Drain hole and kickers
Sub Drain_Hit:bsTrough.addball me : playsoundAtVol"drain" , Drain, 1: End Sub

 'Scoring Rubber
Sub sw1a_hit:vpmTimer.pulseSw 1 : playsoundAtVol"flip_hit_3" , ActiveBall, 1: End Sub
Sub sw1b_hit:vpmTimer.pulseSw 1 : playsoundAtVol"flip_hit_3" , ActiveBall, 1: End Sub
Sub sw1c_hit:vpmTimer.pulseSw 1 : playsoundAtVol"flip_hit_3" , ActiveBall, 1: End Sub

'Star Triggert
Sub SW3_Hit:Controller.Switch(3)=1 : playsoundAtVol"rollover" , ActiveBall, 1: End Sub
Sub SW3_unHit:Controller.Switch(3)=0 : End Sub

'Spinners
Sub sw17_Spin:vpmTimer.PulseSw 17 : playsoundAtVol"fx_spinner" , sw17, 1: End Sub

'Wire Triggers
Sub sw26_Hit:Controller.Switch(26) =1 :playsoundAtVol"rollover" , ActiveBall, 1: End Sub
Sub sw26_UnHit:Controller.Switch(26)=0:End Sub
Sub SW29_Hit:Controller.Switch(29)=1 : playsoundAtVol"rollover" , ActiveBall, 1: End Sub
Sub SW29_unHit:Controller.Switch(29)=0:End Sub
Sub SW30_Hit:Controller.Switch(30)=1 : playsoundAtVol"rollover" , ActiveBall, 1: End Sub
Sub SW30_unHit:Controller.Switch(30)=0:End Sub
Sub SW31_Hit:Controller.Switch(31)=1 : playsoundAtVol"rollover" , ActiveBall, 1: End Sub
Sub SW31_unHit:Controller.Switch(31)=0:End Sub
Sub SW32_Hit:Controller.Switch(32)=1 : playsoundAtVol"rollover" , ActiveBall, 1: End Sub
Sub SW32_unHit:Controller.Switch(32)=0:End Sub
Sub SW33_Hit:Controller.Switch(33)=1 : playsoundAtVol"rollover" , ActiveBall, 1: End Sub
Sub SW33_unHit:Controller.Switch(33)=0:End Sub
Sub SW34_Hit:Controller.Switch(34)=1 : playsoundAtVol"rollover" , ActiveBall, 1: End Sub
Sub SW34_unHit:Controller.Switch(34)=0:End Sub

'Stand Up Targets
Sub sw2_hit:vpmTimer.pulseSw 2 : End Sub
Sub sw25_hit:vpmTimer.pulseSw 25 : End Sub
Sub sw27_hit:vpmTimer.pulseSw 27 : End Sub
Sub sw28_hit:vpmTimer.pulseSw 28 : End Sub

'Bumpers
Sub Bumper1_Hit : vpmTimer.PulseSw(40) : playsoundAtVol SoundFX("fx_bumper1",DOFContactors), ActiveBall, 1: End Sub
Sub Bumper2_Hit : vpmTimer.PulseSw(39) : playsoundAtVol SoundFX("fx_bumper1",DOFContactors), ActiveBall, 1: End Sub
Sub Bumper3_Hit : vpmTimer.PulseSw(38) : playsoundAtVol SoundFX("fx_bumper1",DOFContactors), ActiveBall, 1: End Sub


'-------------------------------------
' Map lights into array
' Set unmapped lamps to Nothing
'-------------------------------------

Set Lights(1)  = L1
Set Lights(2)  = L2
Set Lights(3)  = L3
Set Lights(4)  = L4
Set Lights(5)  = L5
Set Lights(6)  = L6
Set Lights(7)  = LightP01
Set Lights(8)  = LightP05
Set Lights(9)  = LightP09
Set Lights(10) = LightP13
Set Lights(11) = L11  'Light is also Extra Ball Backglass
Set Lights(12) = L12
'Set Lights(13) = 'Ball In Play BackGlass
'Set Lights(14) = ' Player 1 BackGlass
'Set Lights(15) = 'P1 BackGlass
Set Lights(17) = L17
Set Lights(18) = L18
Set Lights(19) = L19
Set Lights(20) = L20
Set Lights(21) = L21
Set Lights(22) = L22
Set Lights(23) = LightP02
Set Lights(24) = LightP06
Set Lights(25) = LightP10
Set Lights(26) = LightP14
'Set Lights(27) = 'Match BackGlass
Set Lights(28) = L28
'Set Lights(29) = 'High Score BackGlass
'Set Lights(30) = 'Player2
'Set Lights(31) = 'P2 BackGlass
Set Lights(33) = L33
Set Lights(34) = L34
Set Lights(35) = L35
Set Lights(36) = L36
Set Lights(37) = L37
Set Lights(38) = L38
Set Lights(39) = LightP03
Set Lights(40) = LightP07
Set Lights(41) = LightP11
Set Lights(42) = LightP15
Lights(43) = array(l43,l43a)
Set Lights(44) = L44
'Set Lights(45) = 'Game Over BackGlass
'Set Lights(46) = 'Player2
'Set Lights(47) = 'P3 BackGlass
Set Lights(49) = L49
Set Lights(50) = L50
Set Lights(51) = L51
Set Lights(52) = L52
Set Lights(53) = L53
Set Lights(54) = L54
Set Lights(55) = LightP04
Set Lights(56) = LightP08
Set Lights(57) = LightP12
Set Lights(58) = L58
Set Lights(59) = L59
Set Lights(60) = L60
'Set Lights(61) = 'Tilt BackGlass
'Set Lights(62) = 'Player 4 BackGlass
'Set Lights(63) = 'P4 BackGlass


'**********************************************************************************************************
'Digital Display
'**********************************************************************************************************

Dim Digits(28)
' 1st Player
Digits(0) = Array(LED10,LED11,LED12,LED13,LED14,LED15,LED16)
Digits(1) = Array(LED20,LED21,LED22,LED23,LED24,LED25,LED26)
Digits(2) = Array(LED30,LED31,LED32,LED33,LED34,LED35,LED36)
Digits(3) = Array(LED40,LED41,LED42,LED43,LED44,LED45,LED46)
Digits(4) = Array(LED50,LED51,LED52,LED53,LED54,LED55,LED56)
Digits(5) = Array(LED60,LED61,LED62,LED63,LED64,LED65,LED66)

' 2nd Player
Digits(6) = Array(LED80,LED81,LED82,LED83,LED84,LED85,LED86)
Digits(7) = Array(LED90,LED91,LED92,LED93,LED94,LED95,LED96)
Digits(8) = Array(LED100,LED101,LED102,LED103,LED104,LED105,LED106)
Digits(9) = Array(LED110,LED111,LED112,LED113,LED114,LED115,LED116)
Digits(10) = Array(LED120,LED121,LED122,LED123,LED124,LED125,LED126)
Digits(11) = Array(LED130,LED131,LED132,LED133,LED134,LED135,LED136)

' 3rd Player
Digits(12) = Array(LED150,LED151,LED152,LED153,LED154,LED155,LED156)
Digits(13) = Array(LED160,LED161,LED162,LED163,LED164,LED165,LED166)
Digits(14) = Array(LED170,LED171,LED172,LED173,LED174,LED175,LED176)
Digits(15) = Array(LED180,LED181,LED182,LED183,LED184,LED185,LED186)
Digits(16) = Array(LED190,LED191,LED192,LED193,LED194,LED195,LED196)
Digits(17) = Array(LED200,LED201,LED202,LED203,LED204,LED205,LED206)

' 4th Player
Digits(18) = Array(LED220,LED221,LED222,LED223,LED224,LED225,LED226)
Digits(19) = Array(LED230,LED231,LED232,LED233,LED234,LED235,LED236)
Digits(20) = Array(LED240,LED241,LED242,LED243,LED244,LED245,LED246)
Digits(21) = Array(LED250,LED251,LED252,LED253,LED254,LED255,LED256)
Digits(22) = Array(LED260,LED261,LED262,LED263,LED264,LED265,LED266)
Digits(23) = Array(LED270,LED271,LED272,LED273,LED274,LED275,LED276)

' Credits
Digits(24) = Array(LED4,LED2,LED6,LED7,LED5,LED1,LED3)
Digits(25) = Array(LED18,LED9,LED27,LED28,LED19,LED8,LED17)
' Balls
Digits(26) = Array(LED39,LED37,LED48,LED49,LED47,LED29,LED38)
Digits(27) = Array(LED67,LED58,LED69,LED77,LED68,LED57,LED59)

Sub DisplayTimer_Timer
  Dim ChgLED,ii,num,chg,stat,obj
  ChgLed = Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
If Not IsEmpty(ChgLED) Then
    If DesktopMode = True Then
    For ii = 0 To UBound(chgLED)
      num = chgLED(ii, 0) : chg = chgLED(ii, 1) : stat = chgLED(ii, 2)
      if (num < 28) then
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


'**********************************************************************************************************
'**********************************************************************************************************

'Bally Eight Ball
'added by Inkochnito
Sub editDips
  Dim vpmDips:Set vpmDips=New cvpmDips
  With vpmDips
    .AddForm 700,400,"Eight Ball - DIP switches"
    .AddChk 7,10,120,Array("Match feature",&H00100000)'dip 21
    .AddChk 205,10,120,Array("Credits displayed",&H00080000)'dip 20
    .AddFrame 2,30,190,"Maximum credits",&H00070000,Array("10 credits",&H00010000,"15 credits",&H00020000,"25 credits",&H00040000,"40 credits",&H00070000)'dip 17&18&19
    .AddFrame 2,106,190,"Right outlane special stays on",&H00200000,Array("only with ball in play",0,"during entire game",&H00200000)'dip 22
    .AddFrame 2,152,190,"Right outlane special award",&H20000000,Array("15,000 points",0,"replay",&H20000000)'dip 30
    .AddFrame 2,198,190,"Bank shot start value",&H00400000,Array("start at 300 points",0,"start at 600 points && 2X",&H00400000)'dip 23
    .AddFrame 2,244,190,"Bank shot extra ball award",&H40000000,Array ("15,000 points",0,"extra ball",&H40000000)'dip 31
    .AddFrame 205,30,190,"High game to date award",&H00000060,Array("no award",0,"1 credit",&H00000020,"2 credits",&H00000040,"3 credits",&H00000060)'dip 6&7
    .AddFrame 205,106,190,"Balls per game",32768,Array("3 balls",0,"5 balls",32768)'dip 16
    .AddFrame 205,152,190,"High score award",&H80000000,Array("extra ball",0,"replay",&H80000000)'dip 32
    .AddFrame 205,198,190,"Melody option",&H00000080,Array("tunes off",0,"tunes on",&H00000080)'dip 8
    .AddFrame 205,244,190,"Top lanes adjust",&H00800000,Array("every lite seperate",0,"1/9 && 4/12 and 2/10 && 3/11 tied",&H00800000)'dip 24
    .AddLabel 50,300,300,20,"After hitting OK, press F3 to reset game with new settings."
    .ViewDips
  End With
End Sub
Set vpmShowDips = GetRef("editDips")

'**********************************************************************************************************
'**********************************************************************************************************
' Start of VPX functions
'**********************************************************************************************************
'**********************************************************************************************************

'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep

Sub RightSlingShot_Slingshot
  vpmTimer.PulseSw 36
    PlaySoundAtVol SoundFX("right_slingshot",DOFContactors), sling1, 1
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.TransZ = -20
    RStep = 0
    RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.TransZ = -10
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.TransZ = 0:RightSlingShot.TimerEnabled = 0:
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
  vpmTimer.PulseSw 37
    PlaySoundAtVol SoundFX("left_slingshot",DOFContactors), sling2, 1
    LSling.Visible = 0
    LSling1.Visible = 1
    sling2.TransZ = -20
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.TransZ = -10
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.TransZ = 0:LeftSlingShot.TimerEnabled = 0:
    End Select
    LStep = LStep + 1
End Sub

'*****************************************
'     FLIPPER SHADOWS
'*****************************************

sub FlipperTimer_Timer()
  FlipperLSh.RotZ = LeftFlipper.currentangle
  FlipperRSh.RotZ = RightFlipper.currentangle

End Sub

'*****************************************
'     BALL SHADOW
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
            BallShadow(b).X = ((BOT(b).X) - (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/7)) '+ 13
        Else
            BallShadow(b).X = ((BOT(b).X) + (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/7)) '- 13
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

' play a looping sound at a location with volume
Sub PlayLoopSoundAtVol(sound, tableobj, Vol)
  PlaySound sound, -1, Vol, AudioPan(tableobj), 0, 0, 1, 0, AudioFade(tableobj)
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
      If BallVel(BOT(b) ) > 1 Then
        rolling(b) = True
        if BOT(b).z < 30 Then ' Ball on playfield
          PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) ), Pan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0, AudioFade(BOT(b) )
        Else ' Ball on raised ramp
          PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) )*.5, Pan(BOT(b) ), 0, Pitch(BOT(b) )+50000, 1, 0, AudioFade(BOT(b) )
        End If
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
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 200, Pan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub

' Thalamus : Exit in a clean and proper way
Sub Table1_exit()
  Controller.Pause = False
  Controller.Stop
End Sub


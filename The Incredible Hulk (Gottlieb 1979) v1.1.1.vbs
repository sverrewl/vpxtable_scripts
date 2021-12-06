Option Explicit
Randomize

' Thalamus 2018-11-01 : Improved directional sounds
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

Const cGameName="hulk",UseSolenoids=1,UseLamps=1,UseGI=0,SSolenoidOn="SolOn",SSolenoidOff="SolOff",SCoin="coin"

LoadVPM"01120100","GTS1.VBS",3.02
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
SolCallback(1)="bsTrough.SolOut"
SolCallback(2)="vpmSolSound SoundFX(""Knocker"",DOFKnocker),"
SolCallback(3)="vpmSolSound SoundFX(""Chime1"", DOFChimes),"
SolCallback(4)="vpmSolSound SoundFX(""Chime2"", DOFChimes),"
SolCallback(5)="vpmSolSound SoundFX(""Chime3"", DOFChimes),"
SolCallback(6)="bsSaucer1.SolOut"
SolCallback(7)="bsSaucer2.SolOut"
SolCallback(8)="dtDrop.SolDropUp"


SolCallBack(17) = "FastFlips.TiltSol"
SolCallback(sLRFlipper) = ""
SolCallback(sLLFlipper) = ""
SolCallback(sURFlipper) = ""
SolCallback(sULFlipper) = ""

Sub SolLFlipper(Enabled)
     If Enabled Then
         PlaySoundAtVol SoundFX("fx_Flipperup",DOFContactors), LeftFlipper, VolFlip:LeftFlipper.RotateToEnd
     Else
         PlaySoundAtVol SoundFX("fx_Flipperdown",DOFContactors), LeftFlipper, VolFlip:LeftFlipper.RotateToStart
     End If
  End Sub

Sub SolRFlipper(Enabled)
     If Enabled Then
         PlaySoundAtVol SoundFX("fx_Flipperup",DOFContactors),RightFlipper,VolFlip:RightFlipper.RotateToEnd
     Else
         PlaySoundAtVol SoundFX("fx_Flipperdown",DOFContactors), RightFlipper, VolFlip:RightFlipper.RotateToStart
     End If
End Sub
'**********************************************************************************************************
'**********************************************************************************************************
'Primitive Flipper Code

Sub FlipperTimer_Timer
  FlipperT1.roty = LeftFlipper.currentangle  + 240
  FlipperT5.roty = RightFlipper.currentangle + 120
End Sub

'*****GI Lights On
dim xx
For each xx in GI:xx.State = 1: Next

'**********************************************************************************************************

'Initiate Table
'**********************************************************************************************************
Dim bsTrough,dtDrop,bsSaucer1,bsSaucer2

Sub Table1_Init
  On Error Resume Next
  With Controller
    .GameName=cGameName
    If Err Then MsgBox"Can't start Game"&cGameName&vbNewLine&Err.Description:Exit Sub
    .SplashInfoLine="The Incredible Hulk Gottlieb" & vbNewLine & "You Suck"
    .HandleMechanics=0
    .ShowDMDOnly=1
    .ShowFrame=0
    .ShowTitle=0
    .Hidden=1
    .Run
    If Err Then MsgBox Err.Description
  End With
  On Error Goto 0

  PinMAMETimer.Interval=PinMAMEInterval
  PinMAMETimer.Enabled=1
  vpmNudge.TiltSwitch=4
  vpmNudge.Sensitivity=3
    vpmNudge.TiltObj=Array(Bumper1,Bumper2)

  Set bsTrough=New cvpmBallStack
  bsTrough.InitSw 0,66,0,0,0,0,0,0
  bsTrough.InitKick BallRelease,90,6
  bsTrough.InitExitSnd SoundFX("ballrelease",DOFContactors), SoundFX("Solenoid",DOFContactors)
  bsTrough.Balls=1

  Set dtDrop=New cvpmDropTarget
  dtDrop.InitDrop Array(sw34,sw32,sw31,sw22,sw21,sw12,sw11),Array(34,32,31,22,21,12,11)
  dtDrop.InitSnd SoundFX("DTDrop",DOFContactors),SoundFX("DTReset",DOFContactors)

  Set bsSaucer1=New cvpmBallstack
  bsSaucer1.InitSaucer Kicker1,41,357,34
  bsSaucer1.KickAngleVar = 1
  bsSaucer1.KickForceVar = 8
  bsSaucer1.InitExitSnd SoundFX("Popper",DOFContactors), SoundFX("Solenoid",DOFContactors)

  Set bsSaucer2=New cvpmBallstack
  bsSaucer2.InitSaucer Kicker2,42,358,34
  bsSaucer2.KickAngleVar = 1.5
  bsSaucer2.KickForceVar = 8
  bsSaucer2.InitExitSnd SoundFX("Popper",DOFContactors), SoundFX("Solenoid",DOFContactors)

End Sub
'**********************************************************************************************************
'**********************************************************************************************************

'Plunger code
'**********************************************************************************************************
Sub Table1_KeyDown(ByVal KeyCode)
  If KeyDownHandler(keycode) Then Exit Sub
  If keycode = PlungerKey Then Plunger.Pullback:playsoundAtVol"plungerpull", Plunger, 1
      If KeyCode = LeftFlipperKey then FastFlips.FlipL True : ' FastFlips.FlipUL True
      If KeyCode = RightFlipperKey then FastFlips.FlipR True :  'FastFlips.FlipUR True
    If keycode=AddCreditKey then playsoundAtVol "coin", Drain, 1: vpmTimer.pulseSW (swCoin1): end if
End Sub

Sub Table1_KeyUp(ByVal KeyCode)
  If KeyUpHandler(keycode) Then Exit Sub
  If keycode = PlungerKey Then Plunger.Fire:PlaySoundAtVol"plunger", Plunger, 1
     If KeyCode = LeftFlipperKey then FastFlips.FlipL False :  'FastFlips.FlipUL False
     If KeyCode = RightFlipperKey then FastFlips.FlipR False :  'FastFlips.FlipUR False
End Sub
'**********************************************************************************************************

'Switches
'**********************************************************************************************************
'**********************************************************************************************************

'Drain
Sub Drain_Hit:PlaySoundAtVol "Drain",drain,1:bsTrough.AddBall Me:End Sub
'Kickers
Sub Kicker1_Hit:bsSaucer1.AddBall 0: playsoundAtVol"popper_ball" , ActiveBall, 1: End Sub
Sub Kicker2_Hit:bsSaucer2.AddBall 0: playsoundAtVol"popper_ball" , ActiveBall, 1: End Sub

'Wire triggers
Sub sw41_Hit:Controller.Switch(44)=1 : playsoundAtVol"rollover" , ActiveBall, VolRol: End Sub
Sub sw41_Unhit:Controller.Switch(44)=0:End Sub
Sub sw60_Hit:Controller.Switch(60)=1 : playsoundAtVol"rollover" , ActiveBall, VolRol: End Sub
Sub sw60_Unhit:Controller.Switch(60)=0:End Sub
Sub sw61_Hit:Controller.Switch(61)=1 : playsoundAtVol"rollover" , ActiveBall, VolRol: End Sub
Sub sw61_Unhit:Controller.Switch(61)=0:End Sub
Sub sw62_Hit:Controller.Switch(62)=1 : playsoundAtVol"rollover" , ActiveBall, VolRol: End Sub
Sub sw62_UnHit:Controller.Switch(62)=0:End Sub
Sub sw64_Hit:Controller.Switch(64)=1 : playsoundAtVol"rollover" , ActiveBall, VolRol: End Sub
Sub sw64_UnHit:Controller.Switch(64)=0:End Sub
Sub sw70_Hit:Controller.Switch(70)=1 : playsoundAtVol"rollover" , ActiveBall, VolRol: End Sub
Sub sw70_UnHit:Controller.Switch(70)=0:End Sub
Sub sw71_Hit:Controller.Switch(71)=1 : playsoundAtVol"rollover" , ActiveBall, VolRol: End Sub
Sub sw71_UnHit:Controller.Switch(71)=0:End Sub
Sub sw72_Hit:Controller.Switch(72)=1 : playsoundAtVol"rollover" , ActiveBall, VolRol: End Sub
Sub sw72_UnHit:Controller.Switch(72)=0:End Sub
Sub sw74_Hit:Controller.Switch(74)=1 : playsoundAtVol"rollover" , ActiveBall, VolRol: End Sub
Sub sw74_UnHit:Controller.Switch(74)=0:End Sub

'Standup Targets
Sub sw10_Hit:vpmTimer.PulseSw 10 : End Sub
Sub sw20_Hit:vpmTimer.PulseSw 20 : End Sub
Sub sw30_Hit:vpmTimer.PulseSw 30 : End Sub

'Drop Target
 Sub Sw34_Dropped: dtDrop.Hit 1: End Sub
 Sub Sw32_Dropped: dtDrop.Hit 2: End Sub
 Sub Sw31_Dropped: dtDrop.Hit 3: End Sub
 Sub Sw22_Dropped: dtDrop.Hit 4: End Sub
 Sub Sw21_Dropped: dtDrop.Hit 5: End Sub
 Sub Sw11_Dropped: dtDrop.Hit 6: End Sub
 Sub Sw12_Dropped: dtDrop.Hit 7: End Sub

 'Spinner
Sub sw24_Spin:vpmTimer.PulseSw 24 : playsoundAtVol"fx_spinner",sw24,VolSpin : End Sub

'Star Rollover
Sub Trigger1_Hit:controller.switch(14)= 1 :playsoundAtVol"rollover" , ActiveBall, 1: End Sub
Sub Trigger1_Unhit:controller.switch(14)= 0 :End Sub
Sub Trigger2_Hit:controller.switch(14)= 1: playsoundAtVol"rollover" , ActiveBall, 1: End Sub
Sub Trigger2_UnHit:controller.switch(14)= 0 :End Sub
Sub Trigger3_Hit:controller.switch(14)= 1: playsoundAtVol"rollover" , ActiveBall, 1: End Sub
Sub Trigger3_UnHit:controller.switch(14)= 0 :End Sub
Sub Trigger4_Hit:controller.switch(14)=1 : playsoundAtVol"rollover" , ActiveBall, 1: End Sub
Sub Trigger4_UnHit:controller.switch(14)= 0 :End Sub

'Scoring Rubbers (5)
 Sub sr1_Hit:vpmTimer.PulseSw 40: playsoundAtVol"rubber_hit_1" , ActiveBall, 1: End Sub
 Sub sr2_Hit:vpmTimer.PulseSw 40: playsoundAtVol"rubber_hit_1" , ActiveBall, 1: End Sub
 Sub sr3_Hit:vpmTimer.PulseSw 40: playsoundAtVol"rubber_hit_1" , ActiveBall, 1: End Sub
 Sub sr4_Hit:vpmTimer.PulseSw 40: playsoundAtVol"rubber_hit_1" , ActiveBall, 1: End Sub
 Sub sr5_Hit:vpmTimer.PulseSw 40: playsoundAtVol"rubber_hit_1" , ActiveBall, 1: End Sub

'Pop Bumpers
Sub Bumper1_Hit:vpmTimer.PulseSw 50 : playsoundAtVol SoundFX("fx_bumper1",DOFContactors), ActiveBall, 1
: B1L1.State = 1:B1L2. State = 1 : Me.TimerEnabled = 1 : DOF 101,2: End Sub
Sub Bumper1_Timer : B1L1.State = 0:B1L2. State = 0 : Me.Timerenabled = 0 : End Sub
Sub Bumper2_Hit:vpmTimer.PulseSw 50 : playsoundAtVol SoundFX("fx_bumper1",DOFContactors), ActiveBall, 1
: B2L1.State = 1:B2L2. State = 1 : Me.TimerEnabled = 1 : DOF 102,2 : End Sub
Sub Bumper2_Timer : B2L1.State = 0:B2L2. State = 0 : Me.Timerenabled = 0 : End Sub

'**********************************************************************************************************

'Map lights to an array
'**********************************************************************************************************
'Set Lights(1)=Light1 'Game Over
'Set Lights(2)=Light2 'Tilt
'Set Lights(3)=Light3 'High Score
Set Lights(4)=Light4
Set Lights(5)=Light5
Set Lights(6)=Light6
Set Lights(7)=Light7
Set Lights(8)=Light8
Set Lights(9)=Light9
Set Lights(10)=Light10
Set Lights(11)=Light11
Set Lights(12)=Light12
Set Lights(13)=Light13
Set Lights(14)=Light14
Set Lights(15)=Light15
Set Lights(16)=Light16
Set Lights(17)=Light17
Set Lights(18)=Light18
Set Lights(19)=Light19
Set Lights(20)=Light20
Set Lights(21)=Light21
Set Lights(22)=Light22
Set Lights(23)=Light23
Set Lights(24)=Light24
Set Lights(25)=Light25
Set Lights(26)=Light26
Set Lights(27)=Light27
Set Lights(28)=Light28
Set Lights(29)=Light29
Set Lights(30)=Light30
Set Lights(31)=Light31
Set Lights(32)=Light32
Set Lights(33)=Light33
Set Lights(34)=Light34
Set Lights(35)=Light35
Set Lights(36)=Light36

'**********************************************************************************************************

'backglass lamps
'**********************************************************************************************************
Dim Digits(32)
Digits(0)=Array(a00,a01,a02,a03,a04,a05,a06,n,a08)
Digits(1)=Array(a10,a11,a12,a13,a14,a15,a16,n,a18)
Digits(2)=Array(a20,a21,a22,a23,a24,a25,a26,n,a28)
Digits(3)=Array(a30,a31,a32,a33,a34,a35,a36,n,a38)
Digits(4)=Array(a40,a41,a42,a43,a44,a45,a46,n,a48)
Digits(5)=Array(a50,a51,a52,a53,a54,a55,a56,n,a58)
Digits(6)=Array(b00,b01,b02,b03,b04,b05,b06,n,b08)
Digits(7)=Array(b10,b11,b12,b13,b14,b15,b16,n,b18)
Digits(8)=Array(b20,b21,b22,b23,b24,b25,b26,n,b28)
Digits(9)=Array(b30,b31,b32,b33,b34,b35,b36,n,b38)
Digits(10)=Array(b40,b41,b42,b43,b44,b45,b46,n,b48)
Digits(11)=Array(b50,b51,b52,b53,b54,b55,b56,n,b58)
Digits(12)=Array(c00,c01,c02,c03,c04,c05,c06,n,c08)
Digits(13)=Array(c10,c11,c12,c13,c14,c15,c16,n,c18)
Digits(14)=Array(c20,c21,c22,c23,c24,c25,c26,n,c28)
Digits(15)=Array(c30,c31,c32,c33,c34,c35,c36,n,c38)
Digits(16)=Array(c40,c41,c42,c43,c44,c45,c46,n,c48)
Digits(17)=Array(c50,c51,c52,c53,c54,c55,c56,n,c58)
Digits(18)=Array(d00,d01,d02,d03,d04,d05,d06,n,d08)
Digits(19)=Array(d10,d11,d12,d13,d14,d15,d16,n,d18)
Digits(20)=Array(d20,d21,d22,d23,d24,d25,d26,n,d28)
Digits(21)=Array(d30,d31,d32,d33,d34,d35,d36,n,d38)
Digits(22)=Array(d40,d41,d42,d43,d44,d45,d46,n,d48)
Digits(23)=Array(d50,d51,d52,d53,d54,d55,d56,n,d58)
Digits(26)=Array(e00,e01,e02,e03,e04,e05,e06,n,e08)
Digits(27)=Array(e10,e11,e12,e13,e14,e15,e16,n,e18)
Digits(24)=Array(f00,f01,f02,f03,f04,f05,f06,n,f08)
Digits(25)=Array(f10,f11,f12,f13,f14,f15,f16,n,f18)



Sub DisplayTimer_Timer
  Dim ChgLED,ii,num,chg,stat,obj
  ChgLed = Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
If Not IsEmpty(ChgLED) Then
    If DesktopMode = True Then
    For ii = 0 To UBound(chgLED)
      num = chgLED(ii, 0) : chg = chgLED(ii, 1) : stat = chgLED(ii, 2)
      if (num < 32) then
        For Each obj In Digits(num)
          If chg And 1 Then obj.State = stat And 1
          chg = chg\2 : stat = stat\2
        Next
      else
        'if char(stat) > "" then msg(num) = char(stat)
      end if
    next
    end if
end if
End Sub



'**********************************************************************************************************

Sub editDips
    Dim vpmDips : Set vpmDips = New cvpmDips
    With vpmDips
      .AddForm 700,400,"System 1 (Multi-Mode sound) - DIP switches"
      .AddFrame 0,0,190,"Coin chute control",&H00040000,Array("seperate",0,"same",&H00040000)'dip 19
      .AddFrame 0,46,190,"Game mode",&H00000400,Array("extra ball",0,"replay",&H00000400)'dip 11
      .AddFrame 0,92,190,"High game to date awards",&H00200000,Array("no award",0,"3 replays",&H00200000)'dip 22
      .AddFrame 0,138,190,"Balls per game",&H00000100,Array("5 balls",0,"3 balls",&H00000100)'dip 9
      .AddFrame 0,184,190,"Tilt effect",&H00000800,Array("game over",0,"ball in play only",&H00000800)'dip 12
      .AddFrame 205,0,190,"Maximum credits",&H00030000,Array("5 credits",0,"8 credits",&H00020000,"10 credits",&H00010000,"15 credits",&H00030000)'dip 17&18
      .AddFrame 205,76,190,"Sound settings",&H80000000,Array("sounds",0,"tones",&H80000000)'dip 32
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


' *******************************************************************************************************
' Positional Sound Playback Functions by DJRobX
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

'*********************************************************************
'                     Supporting Ball & Sound Functions
'*********************************************************************

Function RndNum(min, max)
    RndNum = Int(Rnd() * (max-min + 1) ) + min ' Sets a random number between min and max
End Function

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

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
  Vol = Csng(BallVel(ball) ^2 / VolDiv)
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
    Next
End Sub

'**********************
' Ball Collision Sound
'**********************

Sub OnBallBallCollision(ball1, ball2, velocity)
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / (VolDiv/VolCol), AudioPan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
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
  PlaySound "pinhit_low", 0, Vol(ActiveBall)*VolPi, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Targets_Hit (idx)
  PlaySound "target", 0, Vol(ActiveBall)*VolTarg, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Metals_Thin_Hit (idx)
  PlaySound "metalhit_thin", 0, Vol(ActiveBall)*VolMetal, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Metals_Medium_Hit (idx)
  PlaySound "metalhit_medium", 0, Vol(ActiveBall)*VolMetal, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Metals2_Hit (idx)
  PlaySound "metalhit2", 0, Vol(ActiveBall)*VolMetal, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Gates_Hit (idx)
  PlaySound "gate4", 0, Vol(ActiveBall)*VolGates, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Rubbers_Hit(idx)
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 20 then
    PlaySound "fx_rubber2", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
  End if
  If finalspeed >= 6 AND finalspeed <= 20 then
    RandomSoundRubber()
  End If
End Sub

Sub Posts_Hit(idx)
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 16 then
    PlaySound "fx_rubber2", 0, Vol(ActiveBall)*VolPo, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
  End if
  If finalspeed >= 6 AND finalspeed <= 16 then
    RandomSoundRubber()
  End If
End Sub

Sub RandomSoundRubber()
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySound "rubber_hit_1", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
    Case 2 : PlaySound "rubber_hit_2", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
    Case 3 : PlaySound "rubber_hit_3", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
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
    Case 1 : PlaySound "flip_hit_1", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
    Case 2 : PlaySound "flip_hit_2", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
    Case 3 : PlaySound "flip_hit_3", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
  End Select
End Sub
'**********************************************************************************************************
'cFastFlips by nFozzy
'**********************************************************************************************************
dim FastFlips
Set FastFlips = new cFastFlips
with FastFlips
  .CallBackL = "SolLflipper"  'Point these to flipper subs
  .CallBackR = "SolRflipper"  '...
' .CallBackUL = "SolULflipper"'...(upper flippers, if needed)
' .CallBackUR = "SolURflipper"'...
  .TiltObjects = True 'Optional, if True calls vpmnudge.solgameon automatically. IF YOU GET A LINE 1 ERROR, DISABLE THIS! (or setup vpmNudge.TiltObj!)
' .InitDelay "FastFlips", 100     'Optional, if > 0 adds some compensation for solenoid jitter (occasional problem on Bram Stoker's Dracula)
' .DebugOn = False    'Debug, always-on flippers. Call FastFlips.DebugOn True or False in debugger to enable/disable.
end with

Class cFastFlips
  Public TiltObjects, DebugOn
  Private SubL, SubUL, SubR, SubUR, FlippersEnabled, Delay, LagCompensation, Name

  Private Sub Class_Initialize()
    Delay = 0 : FlippersEnabled = False : DebugOn = False : LagCompensation = False
  End Sub

  'set callbacks
  Public Property Let CallBackL(aInput)  : Set SubL  = GetRef(aInput) : End Property
  Public Property Let CallBackUL(aInput) : Set SubUL = GetRef(aInput) : End Property
  Public Property Let CallBackR(aInput)  : Set SubR  = GetRef(aInput) : End Property
  Public Property Let CallBackUR(aInput) : Set SubUR = GetRef(aInput) : End Property
  Public Sub InitDelay(aName, aDelay) : Name = aName : delay = aDelay : End Sub 'Create Delay

  'call callbacks
  Public Sub FlipL(aEnabled)
    if not FlippersEnabled and not DebugOn then Exit Sub
    subL aEnabled
  End Sub

  Public Sub FlipR(aEnabled)
    if not FlippersEnabled and not DebugOn then Exit Sub
    subR aEnabled
  End Sub

  Public Sub FlipUL(aEnabled)
    if not FlippersEnabled and not DebugOn then Exit Sub
    subUL aEnabled
  End Sub

  Public Sub FlipUR(aEnabled)
    if not FlippersEnabled and not DebugOn then Exit Sub
    subUR aEnabled
  End Sub

  Public Sub TiltSol(aEnabled)  'Handle solenoid / Delay (if delayinit)
    if delay > 0 and not aEnabled then  'handle delay
      vpmtimer.addtimer Delay, Name & ".FireDelay" & "'"
      LagCompensation = True
    else
      if Delay > 0 then LagCompensation = False
      EnableFlippers(aEnabled)
    end if
  End Sub

  Sub FireDelay() : if LagCompensation then EnableFlippers False End If : End Sub

  Private Sub EnableFlippers(aEnabled)
    FlippersEnabled = aEnabled
    if TiltObjects then vpmnudge.solgameon aEnabled
    If Not aEnabled then
      subL False
      subR False
      if not IsEmpty(subUL) then subUL False
      if not IsEmpty(subUR) then subUR False
    End If
  End Sub

End Class

'**********************************************************************************************************
'**********************************************************************************************************

' Thalamus : Exit in a clean and proper way
Sub Table1_exit()
  Controller.Pause = False
  Controller.Stop
End Sub

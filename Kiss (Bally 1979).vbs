Option Explicit
Randomize

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

' Thalamus 2018-07-24
' Table has its own "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Changed UseSolenoids=1 to 2
' No special SSF tweaks yet.
' Is valid for https://vpinball.com/VPBdownloads/kiss-bally-1979-2-2-vpx/ version 2.2

'********************* Rom Selection ***************************
Const SCoin="coin"
 'Const cGameName = "kiss"
 'Const cGameName = "kissb" ,UseSolenoids=1,UseLamps=1,UseGI=0,SSolenoidOn="SolOn",SSolenoidOff="SolOff",SFlipperOn="fx_Flipperup",SFlipperOff="fx_Flipperdown"
 'Const cGameName = "kissp" ' voice prototype currently doesn't work in vpm
 Const cGameName = "kissc" ,UseSolenoids=2,UseLamps=1,UseGI=0,SSolenoidOn="SolOn",SSolenoidOff="SolOff",SFlipperOn="fx_Flipperup",SFlipperOff="fx_Flipperdown"
 'Const cGameName = "kissd"

Const Ballmass = 1.65
Const Ballsize = 49.5

LoadVPM "01550000", "Bally.vbs", 3.26

Dim DesktopMode: DesktopMode = Table1.ShowDT

If DesktopMode = True Then 'Show Desktop components
Ramp1.visible = 1
Ramp15.visible = 1
Else
Ramp1.visible = 0
Ramp15.visible = 0
end If


'Solenoid Call backs
'**********************************************************************************************************
 SolCallback(6) = "vpmSolSound SoundFX(""knocker"",DOFKnocker),"
 SolCallback(7) = "bsTrough.SolOut"
 SolCallback(14) = "DTraised" '"dtbank1.SolDropUp"
 'SolCallback(17) = "vpmSolDiverter Gate,False,Not"
SolCallback(17) = "SolGateDiverter"
 SolCallback(19) = "vpmNudge.SolGameOn"

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub DTraised(enabled)
  if enabled then DTreset.enabled=True
End Sub

Sub DTreset_timer
  dtbank1.DropSol_On
  lightdt1.state = 0
  lightdt2.state = 0
  lightdt3.state = 0
  lightdt4.state = 0
  lightdt5.state = 0
  lightdt6.state = 0
  DTreset.enabled=False
End Sub

Sub SolLFlipper(Enabled)
     If Enabled Then
         PlaySoundAtVol SoundFx("LeftFlipper",DOFContactors),LeftFlipper,.1:LeftFlipper.RotateToEnd
     Else
         PlaySoundAtVol SoundFx("LeftFlipperdown",DOFContactors),LeftFlipper,.1:LeftFlipper.RotateToStart
     End If
  End Sub

Sub SolRFlipper(Enabled)
     If Enabled Then
         PlaySoundAtVol SoundFx("RightFlipper",DOFContactors),RightFlipper,.1:RightFlipper.RotateToEnd
     Else
         PlaySoundAtVol SoundFx("RightFlipperdown",DOFContactors),RightFlipper,.1:RightFlipper.RotateToStart
     End If
End Sub

Sub SolGateDiverter(enabled)
  If Enabled Then
         gate.RotateToStart:PlaySound "metalhit_medium",0,2,.5,0,0,0,0,.8
     Else
         gate.RotateToEnd:PlaySound "metalhit_medium",0,2,.5,0,0,0,0,.8
     End If
End Sub


'**********************************************************************************************************

 'Solenoid Controlled toys
'**********************************************************************************************************

'Primitive Flipper Code
Sub FlipperTimer_Timer
  PrimitiveGate.roty = gate.currentangle
End Sub


'**********************************************************************************************************

'Initiate Table
'**********************************************************************************************************
  Dim bsTrough, dtBank1

Sub Table1_Init
  vpmInit Me
  On Error Resume Next
    With Controller
    .GameName = cGameName
    If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
    .SplashInfoLine = "Kiss - Bally 1979"&chr(13)&"Enjoy"
    .HandleMechanics=0
    .HandleKeyboard=0
    .ShowDMDOnly=1
    .ShowFrame=0
    .ShowTitle=0
        .Hidden = 1
    If Err Then MsgBox Err.Description
  End With
  On Error Goto 0
    Controller.SolMask(0)=0
      vpmTimer.AddTimer 2000,"Controller.SolMask(0)=&Hffffffff'" 'ignore all solenoids - then add the timer to renable all the solenoids after 2 seconds
    Controller.Run
  If Err Then MsgBox Err.Description
  On Error Goto 0

     PinMAMETimer.Interval = PinMAMEInterval
     PinMAMETimer.Enabled = 1

     vpmNudge.TiltSwitch = swTilt
     vpmNudge.Sensitivity = 1
     vpmNudge.TiltObj = Array(Bumper1, Bumper2, Bumper3, Bumper4, LeftSlingshot, RightSlingshot)

     Set bsTrough = New cvpmBallStack
         bsTrough.InitSw 0, 8, 0, 0, 0, 0, 0, 0
         bsTrough.InitKick BallRelease, 80, 6
         bsTrough.InitExitSnd SoundFX("ballrelease",DOFContactors), SoundFX("Solenoid",DOFContactors)
         bsTrough.Balls = 1

     set dtbank1 = new cvpmdroptarget
         dtbank1.initdrop array(sw1, sw2, sw3, sw4), array(1, 2, 3, 4)
         dtbank1.initsnd SoundFX("DTDrop",DOFContactors),SoundFX("DTReset",DOFContactors)

 End Sub

'**********************************************************************************************************
'Plunger code
'**********************************************************************************************************

Sub Table1_KeyDown(ByVal KeyCode)
  If KeyDownHandler(keycode) Then Exit Sub
  If keycode = PlungerKey Then Plunger.Pullback:PlaySoundAt "plungerpull",Plunger
End Sub

Sub Table1_KeyUp(ByVal KeyCode)
  If KeyUpHandler(keycode) Then Exit Sub
  If keycode = PlungerKey Then Plunger.Fire:PlaySoundAt "plunger",Plunger
End Sub

'**********************************************************************************************************
'******************************************************
'           RealTime Updates
'******************************************************

dim defaultEOS,EOSAngle,EOSTorque
defaulteos = leftflipper.eostorque
EOSAngle = 3
EOSTorque = 0.9

Sub GameTimer_Timer()
  RollingSoundUpdate
  BallShadowUpdate

  FlipperLSh.RotZ = LeftFlipper.currentangle
  FlipperRSh.RotZ = RightFlipper.currentangle

  If LeftFlipper.CurrentAngle < LeftFlipper.EndAngle + EOSAngle Then
    LeftFlipper.eostorque = EOSTorque
  Else
    LeftFlipper.eostorque = defaultEOS
  End If

  If RightFlipper.CurrentAngle > RightFlipper.EndAngle - EOSAngle Then
    RightFlipper.eostorque = EOSTorque
  Else
    RightFlipper.eostorque = defaultEOS
  End If



End Sub
'*******************************************************

' Drain holes
 Sub Drain_Hit:PlaySoundAtVol "drain",drain,.5:bsTrough.AddBall Me:End Sub

'scoring rubbers
 Sub Sling3_Hit:vpmTimer.PulseSw 25:PlaySoundAt "Rubber",ActiveBall:End Sub
 Sub Sling4_Hit:vpmTimer.PulseSw 25:PlaySoundAt "Rubber",ActiveBall:End Sub
 Sub Sling5_Hit:vpmTimer.PulseSw 25:PlaySoundAt "Rubber",ActiveBall:End Sub

 'Bumpers
 Sub Bumper1_Hit:vpmTimer.PulseSw 38: PlaySoundAt SoundFX("fx_bumper1",DOFContactors),Bumper1: End Sub
 Sub Bumper2_Hit:vpmTimer.PulseSw 40: PlaySoundAt SoundFX("fx_bumper2",DOFContactors),Bumper2: End Sub
 Sub Bumper3_Hit:vpmTimer.PulseSw 37: PlaySoundAt SoundFX("fx_bumper3",DOFContactors),Bumper3: End Sub
 Sub Bumper4_Hit:vpmTimer.PulseSw 39: PlaySoundAt SoundFX("fx_bumper4",DOFContactors),Bumper4: End Sub

 'Rollovers
 Sub sw34_Hit:Controller.Switch(34) = 1 : PlaySoundAtVol "rollover",sw34,.5: End Sub
 Sub sw34_UnHit:Controller.Switch(34) = 0:End Sub
 Sub sw27_Hit:Controller.Switch(27) = 1 : PlaySoundAtVol "rollover",sw27,.5: End Sub
 Sub sw27_UnHit:Controller.Switch(27) = 0:End Sub
 Sub sw26_Hit:Controller.Switch(26) = 1 : PlaySoundAtVol "rollover",sw26,.5: End Sub
 Sub sw26_UnHit:Controller.Switch(26) = 0:End Sub
 Sub sw33_Hit:Controller.Switch(33) = 1 : PlaySoundAtVol "rollover",sw33,.5: End Sub
 Sub sw33_UnHit:Controller.Switch(33) = 0:End Sub
 Sub sw24_Hit:Controller.Switch(24) = 1 : PlaySoundAtVol "rollover",sw24,.5: End Sub
 Sub sw24_UnHit:Controller.Switch(24) = 0:End Sub
 Sub sw23_Hit:Controller.Switch(23) = 1 : PlaySoundAtVol "rollover",sw23,.5: End Sub
 Sub sw23_UnHit:Controller.Switch(23) = 0:End Sub
 Sub sw5_Hit:Controller.Switch(5) = 1 : PlaySoundAtVol "rollover",sw5,.5: End Sub
 Sub sw5_UnHit:Controller.Switch(5) = 0:End Sub
 Sub sw22_Hit:Controller.Switch(22) = 1 : PlaySoundAtVol "rollover",sw22,.5: End Sub
 Sub sw22_UnHit:Controller.Switch(22) = 0:End Sub
 Sub sw21_Hit:Controller.Switch(21) = 1 : PlaySoundAtVol "rollover",sw21,.5: End Sub
 Sub sw21_UnHit:Controller.Switch(21) = 0:End Sub

 'Spinners
 Sub SPinner1_Spin():vpmTimer.PulseSw 31 : PlaySoundAtVol "fx_spinner",Spinner1,2: End Sub
 Sub SPinner2_Spin():vpmTimer.PulseSw 30 : PlaySoundAtVol "fx_spinner2",Spinner2,2 : End Sub

 ' Droptargets
 Sub sw1_Hit:dtbank1.Hit 1:lightdt1.state = 1:End Sub
 Sub sw2_Hit:dtbank1.Hit 2:lightdt2.state = 1:lightdt5.state = 1:End Sub
 Sub sw3_Hit:dtbank1.Hit 3:lightdt3.state = 1:lightdt6.state = 1:End Sub
 Sub sw4_Hit:dtbank1.Hit 4:lightdt4.state = 1:End Sub

 ' Targets
 Sub sw12_Hit:vpmTimer.PulseSw 12:End Sub
 Sub sw13_Hit:vpmTimer.PulseSw 13:End Sub
 Sub sw14_Hit:vpmTimer.PulseSw 14:End Sub
 Sub sw15_Hit:vpmTimer.PulseSw 15:End Sub
 Sub sw17_Hit:vpmTimer.PulseSw 17:End Sub
 Sub sw18_Hit:vpmTimer.PulseSw 18:End Sub
 Sub sw19_Hit:vpmTimer.PulseSw 19:End Sub
 Sub sw20_Hit:vpmTimer.PulseSw 20:End Sub

'Lights mapped to an array
Set Lights(1)= l1
Set Lights(2)= l2
Set Lights(3)= l3
Set Lights(4)= l4
Set Lights(5)= l5
Set Lights(6)= l6
Set Lights(7)= l7
Set Lights(8)= l8
Lights(9)= array(l9,l9a)
Set Lights(10)= l10
Set Lights(12)= l12
Set Lights(15)= l15
Set Lights(17)= l17
Set Lights(18)= l18
Set Lights(19)= l19
Set Lights(20)= l20
Set Lights(21)= l21
Set Lights(22)= l22
Set Lights(23)= l23
Set Lights(24)= l24
Lights(25)= array(l25,l25a)
Lights(26)= array(Light12,Light21) 'Bumper Lights
Set Lights(28)= l28
Set Lights(30)= l30
Set Lights(31)= l31
Set Lights(33)= l33
Set Lights(34)= l34
Set Lights(35)= l35
Set Lights(36)= l36
Set Lights(37)= l37
Set Lights(38)= l38
Set Lights(39)= l39
Set Lights(40)= l40
Lights(41)= array(l41,l41a)
Lights(42)= array(Light15,Light16) 'Bumper Lights
Set Lights(43)= l43
Set Lights(44)= l44
Set Lights(46)= l46
Set Lights(47)= l47
Set Lights(49)= l49
Set Lights(50)= l50
Set Lights(51)= l51
Set Lights(52)= l52
Set Lights(53)= l53
Set Lights(54)= l54
Set Lights(55)= l55
Set Lights(56)= l56
Lights(57)= array(l57,l57a)
Set Lights(58)= l58
Set Lights(59)= l59
Set Lights(60)= l60
Set Lights(62)= l62
Set Lights(63)= l63

 '************************************
 '          LEDs Display
 'Based on Scapino's 7 digit Reel LEDs
 '************************************


 LampTimer.Interval = 35
 LampTimer.Enabled = 1

 Sub LampTimer_Timer()
     UpdateLeds
 End Sub


 Dim Digits(32)
 Dim Patterns(11)
 Dim Patterns2(11)

 Patterns(0) = 0     'empty
 Patterns(1) = 63    '0
 Patterns(2) = 6     '1
 Patterns(3) = 91    '2
 Patterns(4) = 79    '3
 Patterns(5) = 102   '4
 Patterns(6) = 109   '5
 Patterns(7) = 125   '6
 Patterns(8) = 7     '7
 Patterns(9) = 127   '8
 Patterns(10) = 111  '9

 Patterns2(0) = 128  'empty
 Patterns2(1) = 191  '0
 Patterns2(2) = 134  '1
 Patterns2(3) = 219  '2
 Patterns2(4) = 207  '3
 Patterns2(5) = 230  '4
 Patterns2(6) = 237  '5
 Patterns2(7) = 253  '6
 Patterns2(8) = 135  '7
 Patterns2(9) = 255  '8
 Patterns2(10) = 239 '9

 Set Digits(0) = a0
 Set Digits(1) = a1
 Set Digits(2) = a2
 Set Digits(3) = a3
 Set Digits(4) = a4
 Set Digits(5) = a5
 Set Digits(6) = a6

 Set Digits(7) = b0
 Set Digits(8) = b1
 Set Digits(9) = b2
 Set Digits(10) = b3
 Set Digits(11) = b4
 Set Digits(12) = b5
 Set Digits(13) = b6

 Set Digits(14) = c0
 Set Digits(15) = c1
 Set Digits(16) = c2
 Set Digits(17) = c3
 Set Digits(18) = c4
 Set Digits(19) = c5
 Set Digits(20) = c6

 Set Digits(21) = d0
 Set Digits(22) = d1
 Set Digits(23) = d2
 Set Digits(24) = d3
 Set Digits(25) = d4
 Set Digits(26) = d5
 Set Digits(27) = d6

 Set Digits(28) = e0
 Set Digits(29) = e1
 Set Digits(30) = e2
 Set Digits(31) = e3

 Sub UpdateLeds
     On Error Resume Next
     Dim ChgLED, ii, jj, chg, stat
     ChgLED = Controller.ChangedLEDs(&HFF, &HFFFF)
     If Not IsEmpty(ChgLED)Then
         For ii = 0 To UBound(ChgLED)
             chg = chgLED(ii, 1):stat = chgLED(ii, 2)

             For jj = 0 to 10
                 If stat = Patterns(jj)OR stat = Patterns2(jj)then Digits(chgLED(ii, 0)).SetValue jj
             Next
         Next
     End IF
 End Sub
  Sub NFadeT(nr, a, b)
     Select Case controller.lamp(nr)
         Case False:a.Text = ""
         Case True:a.Text = b
     End Select
End Sub

dim zz
If Table1.ShowDT = false then
    For each zz in DT:zz.Visible = false: Next
else
    For each zz in DT:zz.Visible = true: Next
End If


'**********************************************************************************************************
'**********************************************************************************************************

 'Bally Kiss
 'added by Inkochnito
 Sub editDips
     Dim vpmDips:Set vpmDips = New cvpmDips
     With vpmDips
         .AddForm 700, 400, "Kiss - DIP switches"
         .AddChk 2, 10, 180, Array("Match feature", &H00100000)                                                                                                                                                                                                           'dip 21
         .AddChk 205, 10, 115, Array("Credits display", &H00080000)                                                                                                                                                                                                       'dip 20
         .AddFrame 2, 30, 190, "Maximum credits", &H00070000, Array("5 credits", 0, "10 credits", &H00010000, "15 credits", &H00020000, "20 credits", &H00030000, "25 credits", &H00040000, "30 credits", &H00050000, "35 credits", &H00060000, "40 credits", &H00070000) 'dip 17&18&19
         .AddFrame 2, 160, 190, "Sound features", &H80000080, Array("chime effects", &H80000000, "chime and tunes", 0, "noise", &H00000080, "noises and tunes", &H80000080)                                                                                               'dip 8&32
         .AddFrame 2, 235, 190, "High score to date", &H00000060, Array("no award", 0, "1 credit", &H00000020, "2 credits", &H00000040, "3 credits", &H00000060)                                                                                                          'dip 6&7
         .AddFrame 2, 310, 190, "High score feature", &H00006000, Array("no award", 0, "extra ball", &H00004000, "replay", &H00006000)                                                                                                                                    'dip 14&15
         .AddFrame 205, 30, 190, "Balls per game", 32768, Array("3 balls", 0, "5 balls", 32768)                                                                                                                                                                           'dip 16
         .AddFrame 205, 76, 190, "After completing KISS card 3 times", &H00200000, Array("any letter made is not held over", 0, "any letter made is held over", &H00200000)                                                                                               'dip 22
         .AddFrame 205, 122, 190, "Light-a-line lite", &H00400000, Array("goes on and off", 0, "stays lit", &H00400000)                                                                                                                                                   'dip 23
         .AddFrame 205, 168, 190, "KISS special lites", &H00800000, Array("after 'colossal' lite", 0, "with 'colossal' lite", &H00800000)                                                                                                                                 'dip 24
         .AddFrame 205, 214, 190, "'Opens gate when lit' lite", &H10000000, Array("remains lit", 0, "lites 1 in 3", &H10000000)                                                                                                                                           'dip 29
         .AddFrame 205, 260, 190, "Light-a-line lite", &H20000000, Array("lites for next ball", 0, "comes up same as last ball", &H20000000)                                                                                                                              'dip 30
         .AddFrame 205, 306, 190, "Any A-B-C-D made is", &H40000000, Array("not held in memory", 0, "held in memory for next ball", &H40000000)                                                                                                                           'dip 31
         .AddLabel 50, 382, 300, 20, "After hitting OK, press F3 to reset game with new settings."
         .ViewDips
     End With
 End Sub
 Set vpmShowDips = GetRef("editDips")

'**********************************************************************************************************
'**********************************************************************************************************
'Start of VPX call Backs

'**********************************************************************************************************
'**********************************************************************************************************

'*****GI Lights On
dim xx
For each xx in GI:xx.State = 1: Next

'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep

Sub RightSlingShot_Slingshot
  vpmTimer.PulseSw 35
   PlaySoundAt SoundFX("left_slingshot",DOFContactors),SLING1
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.TransZ = -15
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
  vpmTimer.PulseSw 36
    PlaySoundAt SoundFX("right_slingshot",DOFContactors),SLING2
    LSling.Visible = 0
    LSling1.Visible = 1
    sling2.TransZ = -15
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



' *********************************************************************
'                      Supporting Ball & Sound Functions
' *********************************************************************

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
    Vol = Csng(BallVel(ball) ^2 / 100)
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

function AudioFade(ball)
    Dim tmp
    tmp = ball.y * 2 / Table1.height-1
    If tmp > 0 Then
        AudioFade = Csng(tmp ^10)
    Else
        AudioFade = Csng(-((- tmp) ^10) )
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
  PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 200, Pan(ball1), 0, Pitch(ball1), 0, 0
End Sub

'**************************************************************************
'                 Positional Sound Playback Functions by DJRobX
'**************************************************************************

'Set position as table object (Use object or light but NOT wall) and Vol to 1

Sub PlaySoundAt(sound, tableobj)
    PlaySound sound, 1, 1, Pan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub


'Set all as per ball position & speed.

Sub PlaySoundAtBall(sound)
    PlaySound sound, 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 1, AudioFade(ActiveBall)
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


'**********************************************************************


Sub Pins_Hit (idx):PlaySoundAtBall "pinhit_low":End Sub
Sub Targets_Hit (idx):PlaySoundAt "target",ActiveBall:End Sub
'Sub Metals_Thin_Hit (idx):PlaySoundAtBall "metalhit_thin":End Sub
Sub Metals_Medium_Hit (idx):PlaySoundAtBall "metalhit_medium":End Sub
Sub Metals2_Hit (idx):PlaySoundAtBall "metalhit2":End Sub
Sub Trigger1_Hit:PlaySoundAt "gate",Gate2:End Sub

' Gates
Sub Gate1_Hit():PlaySoundAt "gate4",Gate1:End Sub
Sub Gate2_Hit():PlaySoundAt "gate",Gate2:End Sub
Sub Gate3_Hit():PlaySoundAt "gate4",Gate3:End Sub


'***********************
' Random Rubber Bumps
'***********************


Sub Rubbers_Hit(idx)
    PlaySoundAtBallVol "rubber_hit_" & Int(Rnd*3)+1, .1
End Sub

Sub Posts_Hit(idx)
    PlaySoundAtBallVol "rubber_hit_" & Int(Rnd*3)+1, .1
End Sub


'***********************
' Random Flipper Hits
'***********************

Sub LeftFlipper_Collide(parm)
    PlaySoundAtBallVol "flip_hit_" & Int(Rnd*3)+1, 1
End Sub

Sub RightFlipper_Collide(parm)
    PlaySoundAtBallVol "flip_hit_" & Int(Rnd*3)+1, 1
End Sub

'*********** BALL SHADOW *********************************t
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

Dim NextOrbitHit:NextOrbitHit = 0

Sub Wall40_Hit()
  if BallVel(ActiveBall) > .3 and Timer > NextOrbitHit then
    RandomBump .2, 50000
    ' Schedule the next possible sound time.  This prevents it from rapid-firing noises too much.
    ' Lowering these numbers allow more closely-spaced clunks.
    NextOrbitHit = Timer + .3 + (Rnd * .2)
  end if
End Sub

Sub Metals_Thin_Hit(idx)
  if BallVel(ActiveBall) > .05 and Timer > NextOrbitHit then
    RandomBumpMetals_Thin 1, 2000
    ' Schedule the next possible sound time.  This prevents it from rapid-firing noises too much.
    ' Lowering these numbers allow more closely-spaced clunks.
    NextOrbitHit = Timer + .1 + (Rnd * .2)
  end if
End Sub

Sub RandomBump(voladj, freq)
  dim BumpSnd:BumpSnd= "fx_rampbump" & CStr(Int(Rnd*7)+1)
  If Table1.VersionMinor > 3 OR Table1.VersionMajor > 10 Then
    PlaySound BumpSnd, 0, Vol(ActiveBall)*voladj, Pan(ActiveBall), 0, freq, 0, 1, AudioFade(ActiveBall)
  Else
    PlaySound BumpSnd, 0, Vol(ActiveBall)*voladj, Pan(ActiveBall), 0, freq, 0, 1
  End If
End Sub

Sub RandomBumpMetals_Thin(voladj, freq)
  dim BumpSnd:BumpSnd= "fx_sensor" & CStr(Int(Rnd*3)+1)
  If Table1.VersionMinor > 3 OR Table1.VersionMajor > 10 Then
    PlaySound BumpSnd, 0, Vol(ActiveBall)*voladj, Pan(ActiveBall), 0, freq, 0, 1, AudioFade(ActiveBall)
  Else
    PlaySound BumpSnd, 0, Vol(ActiveBall)*voladj, Pan(ActiveBall), 0, freq, 0, 1
  End If
End Sub

Sub Table1_Exit()

End Sub

' Thalamus : Exit in a clean and proper way
Sub Table1_exit()
  Controller.Pause = False
  Controller.Stop
End Sub


Option Explicit
Randomize

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Const cGameName="rflshdlx",UseSolenoids=2,UseLamps=1,UseGI=0,SSolenoidOn="SolOn",SSolenoidOff="SolOff",SFlipperOn="Flipper2",SFlipperOff="Flipper"
Const SCoin="coin3",cCredits=""

Const Ballsize = 52


' Thalamus 2020 January : Improved directional sounds
' !! NOTE : Table not verified yet !!

' Options
' Volume devided by - lower gets higher sound

Const VolDiv = 2000    ' Lower number, louder ballrolling/collition sound
Const VolCol = 10      ' Ball collition divider ( voldiv/volcol )

LoadVPM "01560000","sys80.vbs",2.0

If Table1.ShowDT = True Then 'Show Desktop components
DisplayTimer.Enabled = 1
Else
DisplayTimer.Enabled = 0
End if



Set LampCallback=GetRef("UpdateMultipleLamps")


gi_2.state=1
gi_3.state=1
gi_8.state=1
gi_9.state=1
gi_10.state=1
gi_11.state=1
GI_12.state=1
gi_14.state=1
gi_16.state=1
gi_17.state=1
gi_19.state=1
gi_21.state=1
GI_24.state=1
GI_28.state=1
GI_30.state=1
GI_33.state=1
GI_38.state=1
GI_39.state=1
GI_40.state=1
GI_41.state=1
GI_46.state=1
GI_52.state=1
GI_53.state=1

'*************************************************************
'Solenoid Call backs
'**********************************************************************************************************




Const sbkHole=5
Const sdtT=6


Const sKnocker=8
Const sOutHole=9

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"
SolCallback(sbkHole)="bsHole.SolOut"
SolCallback(sdtT)="dtT.SolDropUp"

SolCallback(sKnocker)="vpmSolSound""knocker"","
SolCallback(sOutHole)="bsTrough.SolOut"


Sub sollFlipper(Enabled)
     If Enabled Then
         PlaySoundAtVol SoundFX("fx_Flipperup",DOFContactors), LeftFlipper, 1:LeftFlipper.RotateToEnd
     Else
         PlaySoundAtVol SoundFX("fx_Flipperdown",DOFContactors), LeftFlipper, 1:LeftFlipper.RotateToStart
     End If
  End Sub

Sub solrflipper(Enabled)
     If Enabled Then
         PlaySoundAtVol SoundFX("fx_Flipperup",DOFContactors), RightFlipper, 1:RightFlipper.RotateToEnd
     Else
         PlaySoundAtVol SoundFX("fx_Flipperdown",DOFContactors), RightFlipper, 1:RightFlipper.RotateToStart
     End If
End Sub

'**********************************************************************************************************

'Solenoid Controlled toys
'**********************************************************************************************************


'Primitive Flipper
Sub FlipperTimer_Timer
  FlipperT1.roty = LeftFlipper.currentangle  + 237
  FlipperT5.roty = RightFlipper.currentangle + 120


End Sub

Dim VarRol, VarHidden

Sub UpdateMultipleLamps
dim dropxx

End Sub





'**********************************************************************************************************

'Initiate Table
'**********************************************************************************************************

Dim bsTrough, bsKicker, bsHole, DTT, DTL, DTR, dtcl, DTcR,bssaucer

Sub Table1_Init
  'vpmInit Me
  On Error Resume Next
    With Controller
    .GameName = cGameName
    If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
    .SplashInfoLine = "Royal Flush Deluxe"
    .HandleMechanics=0
    .HandleKeyboard=0
    .ShowDMDOnly=1
    .ShowFrame=0
    .ShowTitle=0
        .hidden = 1
    If Err Then MsgBox Err.Description
  End With
  On Error Goto 0
    Controller.SolMask(0)=0
      vpmTimer.AddTimer 2000,"Controller.SolMask(0)=&Hffffffff'" 'ignore all solenoids - then add the timer to renable all the solenoids after 2 seconds
    Controller.Run
  If Err Then MsgBox Err.Description
  On Error Goto 0

  PinMAMETimer.Interval=PinMAMEInterval
  PinMAMETimer.Enabled=1

    vpmNudge.TiltSwitch=57
    vpmNudge.Sensitivity=5
  vpmNudge.TiltObj = Array(Bumper1,LeftslingShot,RightslingShot)

    Set bsTrough=New cvpmBallStack
        bsTrough.InitExitSnd  SoundFX("ballrelease",DOFContactors), SoundFX("Solenoid",DOFContactors)
    bsTrough.InitNoTrough newball,67,100,1



    Set bsHole=New cvpmBallStack
        bsHole.InitSaucer hole46,46,200,15
        bsHole.InitExitSnd  SoundFX("Popper",DOFContactors), SoundFX("Solenoid",DOFContactors)
        bsHole.KickForceVar = 3
        bsHole.KickAngleVar = 3

  set DTT = new cvpmDropTarget
    DTT.InitDrop Array(tga9,tga8,tga7,tga6,tga5,tga4,tga3,tga2,tga1),Array(40,50,60,41,51,61,42,52,62)
    DTT.InitSnd SoundFX("",DOFContactors),SoundFX("DTReset",DOFContactors)


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




Sub Drain_Hit:Drain.DestroyBall:bsTrough.AddBall 0: End Sub   'switch 67
Sub hole46_Hit:bsHole.AddBall 0:End Sub     'switch 46
Sub Bumper1_Hit:vpmTimer.PulseSw(65):RandomSoundbumper:End Sub  'switch 65




Sub tga9_Hit():dtT.Hit 1: PlaySoundAt "DTDrop", tga9 : End Sub
Sub tga8_Hit():dtT.Hit 2: PlaySoundAt "DTDrop", tga8 : End Sub
Sub tga7_Hit():dtT.Hit 3: PlaySoundAt "DTDrop", tga7 : End Sub
Sub tga6_Hit():dtT.Hit 4: PlaySoundAt "DTDrop", tga6 : End Sub
Sub tga5_Hit():dtT.Hit 5: PlaySoundAt "DTDrop", tga5 : End Sub
Sub tga4_Hit():dtT.Hit 6: PlaySoundAt "DTDrop", tga4 : End Sub
Sub tga3_Hit():dtT.Hit 7: PlaySoundAt "DTDrop", tga3 : End Sub
Sub tga2_Hit():dtT.Hit 8: PlaySoundAt "DTDrop", tga2 : End Sub
Sub tga1_Hit():dtT.Hit 9: PlaySoundAt "DTDrop", tga1 : End Sub


Sub sw53a_Hit:vpmTimer.PulseSw(53):End Sub
Sub sw43a_Hit:vpmTimer.PulseSw(43):End Sub
Sub sw63a_Hit:vpmTimer.PulseSw(63):End Sub
Sub sw53_Hit:vpmTimer.PulseSw(53):End Sub
Sub sw43_Hit:vpmTimer.PulseSw(43):End Sub
Sub sw63_Hit:vpmTimer.PulseSw(63):End Sub
Sub sw65_Hit:vpmTimer.PulseSw(65):End Sub
Sub sw65a_Hit:vpmTimer.PulseSw(65):End Sub
Sub sw45_Hit:vpmTimer.PulseSw(45):End Sub
Sub sw45a_Hit:vpmTimer.PulseSw(45):End Sub
Sub sw56_Hit:vpmTimer.PulseSw(56):End Sub
Sub sw56b_Hit:vpmTimer.PulseSw(56):End Sub


Sub sw44_Hit:Controller.Switch(44)=1:End Sub
Sub sw44_unHit:Controller.Switch(44)=0:End Sub
Sub sw64_Hit:Controller.Switch(64)=1:End Sub
Sub sw64_unHit:Controller.Switch(64)=0:End Sub
Sub sw44a_Hit:Controller.Switch(44)=1:End Sub
Sub sw44a_unHit:Controller.Switch(44)=0:End Sub
Sub sw64a_Hit:Controller.Switch(64)=1:End Sub
Sub sw64a_unHit:Controller.Switch(64)=0:End Sub

Sub sw55_Hit:Controller.Switch(55)=1:End Sub
Sub sw55_unHit:Controller.Switch(55)=0:End Sub

Sub sw66_Hit:Controller.Switch(66)=1: End Sub
Sub sw66_unHit:Controller.Switch(66)=0:End Sub

Sub sw54_Hit:Controller.Switch(54)=1:End Sub
Sub sw54_unHit:Controller.Switch(54)=0:End Sub

Sub sw67_Hit:Controller.Switch(67)=1:End Sub
Sub sw67_unHit:Controller.Switch(67)=0:End Sub


Dim BallsInPlay


'**********************************************************************************************************
' Map Lights to Array
'**********************************************************************************************************


lights(13)=array(Light13,light13a)
lights(14)=array(Light14,light14a)
lights(15)=array(Light15,light15a)
lights(16)=array(Light16,light16a)
lights(18)=array(Light18,light18a)
lights(20)=array(Light20,light20a)
lights(34)=array(Light34,light34a)
set lights(17)=Light17
set lights(19)=Light19
set lights(21)=Light21
set lights(22)=light22
set lights(23)=light23
set Lights(24)=light24
set lights(25)=light25
set lights(26)=light26
set lights(27)=light27
set lights(28)=light28
set lights(29)=light29
set lights(30)=light30
set lights(31)=light31
set lights(32)=light32
set lights(33)=light33
set lights(35)=light35
set lights(40)=light40
set lights(41)=light41
Set Lights(12)=Light12








'**********************************************************************************************************
' Backglass Light Displays
'**********************************************************************************************************


Sub lampTimer_Timer()

If Table1.ShowDT = True Then 'Show Desktop components
light35a.state=light35.state
light35b.state=light35.state
light40a.state=light40.state
light40b.state=light40.state
light41a.state=light41.state
light41b.state=light41.state
End if
If light12.state=1 Then
        flippergate12.RotateToEnd
     Else
        flippergate12.RotateToStart
     End If
  gate12.roty = flippergate12.currentangle+90
End Sub


Dim Digits(44)
Digits(0)=Array(a00,a01,a02,a03,a04,a05,a06,a07,a08)
Digits(1)=Array(a10,a11,a12,a13,a14,a15,a16,a17,a18)
Digits(2)=Array(a20,a21,a22,a23,a24,a25,a26,a27,a28)
Digits(3)=Array(a30,a31,a32,a33,a34,a35,a36,a37,a38)
Digits(4)=Array(a40,a41,a42,a43,a44,a45,a46,a47,a48)
Digits(5)=Array(a50,a51,a52,a53,a54,a55,a56,a57,a58)
Digits(6)=Array(b00,b01,b02,b03,b04,b05,b06,b07,b08)

Digits(7)=Array(b10,b11,b12,b13,b14,b15,b16,b17,b18)
Digits(8)=Array(b20,b21,b22,b23,b24,b25,b26,b27,b28)
Digits(9)=Array(b30,b31,b32,b33,b34,b35,b36,b37,b38)
Digits(10)=Array(b40,b41,b42,b43,b44,b45,b46,b47,b48)
Digits(11)=Array(b50,b51,b52,b53,b54,b55,b56,b57,b58)
Digits(12)=Array(c00,c01,c02,c03,c04,c05,c06,c07,c08)
Digits(13)=Array(c10,c11,c12,c13,c14,c15,c16,c17,c18)

Digits(14)=Array(c20,c21,c22,c23,c24,c25,c26,c27,c28)
Digits(15)=Array(c30,c31,c32,c33,c34,c35,c36,c37,c38)
Digits(16)=Array(c40,c41,c42,c43,c44,c45,c46,c47,c48)
Digits(17)=Array(c50,c51,c52,c53,c54,c55,c56,c57,c58)
Digits(18)=Array(d00,d01,d02,d03,d04,d05,d06,d07,d08)
Digits(19)=Array(d10,d11,d12,d13,d14,d15,d16,d17,d18)
Digits(20)=Array(d20,d21,d22,d23,d24,d25,d26,d27,d28)

Digits(21)=Array(d30,d31,d32,d33,d34,d35,d36,d37,d38)
Digits(22)=Array(d40,d41,d42,d43,d44,d45,d46,d47,d48)
Digits(23)=Array(d50,d51,d52,d53,d54,d55,d56,d57,d58)
Digits(24)=Array(e00,e01,e02,e03,e04,e05,e06,e07,e08)
Digits(25)=Array(e10,e11,e12,e13,e14,e15,e16,e17,e18)
Digits(26)=Array(f00,f01,f02,f03,f04,f05,f06,f07,f08)
Digits(27)=Array(f10,f11,f12,f13,f14,f15,f16,f17,f18)


Digits(28) = Array(e2,e3,e7,e4,e5,e1,e6,e23)
Digits(29) = Array(e9,e17,e22,e19,e20,e8,e21,e24)
Digits(30) = Array(f2,f3,f7,f4,f5,f1,f6,f23)
Digits(31) = Array(f9,f17,f22,f19,f20,f8,f21,f24)


Sub DisplayTimer_Timer
  Dim ChgLED,ii,num,chg,stat,obj
  ChgLed = Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
If Not IsEmpty(ChgLED) Then

    For ii = 0 To UBound(chgLED)
      num = chgLED(ii, 0) : chg = chgLED(ii, 1) : stat = chgLED(ii, 2)
      if (num < 44) then
        For Each obj In Digits(num)
          If chg And 1 Then obj.State = stat And 1
          chg = chg\2 : stat = stat\2
        Next
      else
        'if char(stat) > "" then msg(num) = char(stat)
      end if
    next

end if
End Sub


' '**********************************************************************************************************
'   'Gottlieb System 80A Sound only (S)board
'  'added by Inkochnito
'  Sub editDips
'     Dim vpmDips:Set vpmDips=New cvpmDips
'     With vpmDips
'        .AddForm 700, 400, "System 80A with Sound only (S)board - DIP switches"
'        .AddFrame 0, 0, 190, "Maximum credits", 49152, Array("8 credits", 0, "10 credits", 32768, "15 credits", &H00004000, "25 credits", 49152)                                                                                   'dip 15&16
'        .AddFrame 0, 76, 190, "Coin chute 1 and 2 control", &H00002000, Array("seperate", 0, "same", &H00002000)                                                                                                                   'dip 14
'        .AddFrame 0, 122, 190, "3rd coin chute credits control", &H20000000, Array("no effect", 0, "add 9", &H20000000)                                                                                                            'dip 30
'        .AddFrame 205, 0, 190, "High score to date awards", &H00C00000, Array("not displayed and no award", 0, "displayed and no award", &H00800000, "displayed and 2 credits", &H00400000, "displayed and 3 credits", &H00C00000) 'dip 23&24
'        .AddFrame 205, 76, 190, "Game mode", &H10000000, Array("replay", 0, "extra ball", &H10000000)                                                                                                                              'dip 29
'        .AddFrame 205, 122, 190, "Playfield special", &H00200000, Array("replay", 0, "extra ball", &H00200000)                                                                                                                     'dip 22
'        .AddFrame 205, 168, 190, "Replay limit", &H04000000, Array("no limit", 0, "one per game", &H04000000)                                                                                                                      'dip 27
'        .AddFrame 205, 214, 190, "Balls per game", &H01000000, Array("5 balls", 0, "3 balls", &H01000000)                                                                                                                          'dip 25
'        .AddFrame 205, 260, 190, "Novelty mode", &H08000000, Array("normal game mode", 0, "50K per special/extra ball", &H08000000)                                                                                                'dip 28
'        .AddChk 0, 170, 190, Array("Match feature", &H02000000)                                                                                                                                                                    'dip 26
'        .AddChk 0, 185, 190, Array("Background sound", &H40000000)                                                                                                                                                                 'dip 31
'        .AddChk 0, 200, 190, Array("Dip 6 (spare)", &H00000020)                                                                                                                                                                    'dip 6
'        .AddChk 0, 215, 190, Array("Dip 7 (spare)", &H00000040)                                                                                                                                                                    'dip 7
'        .AddChk 0, 230, 190, Array("Dip 8 (spare)", &H00000080)                                                                                                                                                                    'dip 8
'        .AddChk 0, 245, 190, Array("Dip 32 (spare or game option)", &H80000000)                                                                                                                                                    'dip 32
'        .AddLabel 50, 310, 300, 20, "After hitting OK, press F3 to reset game with new settings."
'        .ViewDips
'     End With
'  End Sub
'  Set vpmShowDips=GetRef("editDips")

' ==================================================================
' Gottlieb Roal Flush
' originally added by Inkochnito
' Added Coins chute by Mike da Spike
Sub editDips
  Dim vpmDips : Set vpmDips = New cvpmDips
  With vpmDips
        .AddForm  700,400,"Royal Flush - DIP switches"
      .AddFrame 2,4,190,"Left Coin Chute (Coins/Credit)",&H0000001F,Array("1/4",&H00000018,"1/2",&H00000010,"1/1",&H00000000,"2/1",&H0000000A) 'Dip 1-5
      .AddFrame 2,80,190,"Right Coin Chute (Coins/Credit)",&H00001F00,Array("1/4",&H00001800,"1/2",&H00001000,"1/1",&H00000000,"2/1",&H00000A00) 'Dip 9-13
      .AddFrame 2,160,190,"Center Coin Chute (Coins/Credit)",&H001F0000,Array("1/4",&H00180000,"1/2",&H00100000,"1/1",&H00000000,"2/1",&H000A0000) 'Dip 17-21
    .AddFrame 2,240,190,"3rd coin chute credits control",&H20000000,Array("no effect",0,"add 9",&H20000000)'dip 30
    .AddFrame 207,4,190,"Maximum credits",49152,Array("8 credits",0,"10 credits",32768,"15 credits",&H00004000,"25 credits",49152)'dip 15&16
    .AddFrame 207,80,190,"Coin chute 1 and 2 control",&H00002000,Array("seperate",0,"same",&H00002000)'dip 14
    .AddFrame 207,126,190,"Playfield special",&H00200000,Array("replay",0,"extra ball",&H00200000)'dip 22
    .AddChk 207,174,190, Array("Dip 6 (spare)", &H00000020)
    .AddChk 207,189,190, Array("Dip 7 (spare)", &H00000040)
    .AddChk 207,204,190, Array("Dip 8 (spare)", &H00000080)
    .AddChk 207,219,190, Array("Dip 31 (spare or game option)", &H40000000)
    .AddChk 207,234,190, Array("Dip 32 (spare or game option)", &H80000000)
    .AddChk 207,249,180,Array("Match feature",&H02000000)'dip 26
    .AddFrame 412,4,190,"High game to date awards",&H00C00000,Array("not displayed and no award", 0, "displayed and no award", &H00800000, "displayed and 2 credits", &H00400000, "displayed and 3 credits", &H00C00000)'dip 23&24
    .AddFrame 412,80,190,"Balls per game",&H01000000,Array("5 balls",0,"3 balls",&H01000000)'dip 25
    .AddFrame 412,126,190,"Replay limit",&H04000000,Array("no limit",0,"one per game",&H04000000)'dip 27
    .AddFrame 412,172,190,"Novelty",&H08000000,Array("normal",0,"yes",&H08000000)'dip 28
    .AddFrame 412,218,190,"Game mode",&H10000000,Array("replay",0,"extra ball",&H10000000)'dip 29
    .AddLabel 160,300,300,20,"After hitting OK, press F3 to reset game with new settings."
    .ViewDips
  End With
End Sub
Set vpmShowDips = GetRef("editDips")

'**********************************************************************************************************

' *********************************************************************

          'Start of VPX call back Functions

' *********************************************************************
' *********************************************************************


'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep


Sub RightSlingShot_Slingshot
  vpmTimer.PulseSw(56)
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
  vpmTimer.PulseSw(56)
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

Function BallRollVol(ball) ' Calculates the Volume of the sound based on the ball speed
   BallRollVol = Csng(BallVel(ball) ^2 / (80000 - (79900 * Log(RollVol) / Log(100))))
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

Sub rollingmetal_Hit (idx)
  PlaySound "Wire Ramp", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub


Sub Pins_Hit (idx)
  PlaySound "pinhit_low", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub



Sub Metals_Thin_Hit (idx)
  PlaySound "metalhit_thin", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Metals_Medium_Hit (idx)
  PlaySound "metalhit_medium", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub


Sub Gates_Hit (idx)
  PlaySound "gate4", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub


Sub Triggers_Hit (idx)
  PlaySound "fx_sensor", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
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



Sub RandomSoundRubber()
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySound "rubber_hit_1", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 2 : PlaySound "rubber_hit_2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 3 : PlaySound "rubber_hit_3", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End Select
End Sub

Sub RandomSoundbumper()
  Select Case Int(Rnd*5)+1
    Case 1 : PlaySoundAtVol SoundFX("fx_Bumper",DOFContactors), ActiveBall, 1
    Case 2 : PlaySoundAtVol SoundFX("fx_Bumper1",DOFContactors), ActiveBall, 1
    Case 3 : PlaySoundAtVol SoundFX("fx_Bumper2",DOFContactors), ActiveBall, 1
    Case 4 : PlaySoundAtVol SoundFX("fx_bumper3",DOFContactors), ActiveBall, 1
    Case 5 : PlaySoundAtVol SoundFX("fx_bumper4",DOFContactors), ActiveBall, 1
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

' Thalamus - 2021-04-30 : added proper exit

' Thalamus : Exit in a clean and proper way
Sub Table1_exit()
  Controller.Pause = False
  Controller.Stop
End Sub

